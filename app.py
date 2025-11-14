#v4版本：新增主動推播功能與排程器
'''
import os
import json
from datetime import datetime
from dotenv import load_dotenv, set_key # 引入 set_key 來儲存 userId
from flask import Flask, request, abort

# --- LINE Bot SDK ---
from linebot.v3 import (
    WebhookHandler
)
from linebot.v3.exceptions import (
    InvalidSignatureError
)
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    PushMessageRequest, # 【v4 新增】 引入 Push API
    TextMessage
)
from linebot.v3.webhooks import (
    MessageEvent,
    TextMessageContent
)

# --- Google SDKs ---
import google.generativeai as genai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# --- Scheduler ---
from apscheduler.schedulers.background import BackgroundScheduler # 【v4 新增】

# --- 1. 初始設定 ---

# 載入 .env 檔案中的環境變數 (金鑰)
load_dotenv()
# 將 .env 檔案路徑存起來，方便等等寫入
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')

# 建立 Flask 伺服器
app = Flask(__name__)

# --- LINE 設定 ---
LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')
LINE_CHANNEL_ACCESS_TOKEN = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')
line_configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)
# 【v4 新增】 建立一個全域的 API 實例，供排程器使用
global_api_client = ApiClient(line_configuration)
global_line_bot_api = MessagingApi(global_api_client)

# --- 【v4 新增】 讀取要推播的 User ID ---
# 我們將 User ID 儲存在 .env 中，確保伺服器重啟後還在
USER_ID_TO_PUSH = os.getenv('USER_ID_TO_PUSH')

# --- Gemini 設定 ---
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')
genai.configure(api_key=GEMINI_API_KEY)
gemini_model = genai.GenerativeModel('gemini-2.5-flash') # 使用您帳戶可用的模型

# --- Google Sheets 設定 ---
GOOGLE_SHEET_NAME = os.getenv('GOOGLE_SHEET_NAME')
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
client = gspread.authorize(creds)
try:
    spreadsheet = client.open(GOOGLE_SHEET_NAME)
    worksheet = spreadsheet.sheet1
    app.logger.info(f"成功開啟 Google Sheet: {GOOGLE_SHEET_NAME}")
except gspread.exceptions.SpreadsheetNotFound:
    app.logger.error(f"錯誤：找不到名稱為 '{GOOGLE_SHEET_NAME}' 的 Google Sheet")
    exit()

# 在啟動時抓取一次今天日期
app.config['TODAYS_DATE'] = datetime.now().strftime('%Y-%m-%d')


# --- 2. 建立 Webhook 監聽路徑 (與 v1 相同) ---

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        app.logger.error("簽名錯誤")
        abort(400)
    return 'OK'

# --- 3. 處理「文字訊息」事件 (v4 升級版：自動儲存 UserID、手動觸發晨報) ---

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    global USER_ID_TO_PUSH # 宣告我們要使用全域變數
    message_text = event.message.text
    reply_token = event.reply_token
    user_id = event.source.user_id # 【v4 新增】 抓取使用者的 User ID
    reply_text = ""

    # 【v4 新增】 檢查並儲存 User ID
    # 如果 .env 中還沒有 USER_ID_TO_PUSH，我們就自動幫他存起來
    if not USER_ID_TO_PUSH:
        USER_ID_TO_PUSH = user_id
        set_key(dotenv_path, "USER_ID_TO_PUSH", USER_ID_TO_PUSH)
        app.logger.info(f"成功儲存 User ID: {user_id} 到 .env 檔案")
        reply_text = f"你好！我已經將您的 User ID 設為推播目標。\n\n"
    else:
        # 如果已經存過，也檢查一下，確保是同一個人
        if USER_ID_TO_PUSH != user_id:
             app.logger.warning(f"偵測到不同的 User ID: {user_id} (已設定: {USER_ID_TO_PUSH})")


    # 【指令路由器 (Router)】
    
    # 意圖 1：智能寫入 (WRITE)
    if "提醒" in message_text or "新增" in message_text or "指派" in message_text or "待辦" in message_text:
        try:
            reply_text += handle_write_task(message_text)
        except Exception as e:
            app.logger.error(f"智能寫入失敗: {e}")
            reply_text += f"❌ 任務新增失敗：\n{e}"
    
    # 意圖 2：被動查詢 (QUERY)
    elif "查詢" in message_text or "進度" in message_text or "有什麼" in message_text or "幫我找" in message_text:
        try:
            reply_text += handle_query_task(message_text)
        except Exception as e:
            app.logger.error(f"任務查詢失敗: {e}")
            reply_text += f"❌ 任務查詢失敗：\n{e}"

    # 意圖 3：【v4 新增】 手動觸發晨報 (TEST PUSH)
    elif "晨報" in message_text or "summary" in message_text:
        try:
            app.logger.info(f"手動觸發晨報，推播給: {USER_ID_TO_PUSH}")
            # 呼叫主動推播函式
            send_daily_summary()
            reply_text += "✅ 晨報已發送，請檢查您的 LINE 訊息。"
        except Exception as e:
            app.logger.error(f"手動觸發晨報失敗: {e}")
            reply_text += f"❌ 晨報發送失敗：\n{e}"
    
    # 意圖 4：閒聊 (CHAT)
    else:
        reply_text += "您好，我是一個專案管理助理。\n\n- 新增任務：「提醒我...」\n- 查詢任務：「查詢...」\n- 測試晨報：「晨報」"


    # 統一回覆
    global_line_bot_api.reply_message_with_http_info(
        ReplyMessageRequest(
            reply_token=reply_token,
            messages=[TextMessage(text=reply_text.strip())]
        )
    )

# --- 4. 核心功能：智能寫入 (Handle Write Task) ---
# (此函式與 v3 完全相同，故折疊)
def handle_write_task(message_text):
    app.logger.info(f"執行智能寫入：{message_text}")
    prompt = f"""
    你是一個專案管理 AI 助理。請從以下使用者訊息中，解析並提取資訊，並嚴格按照指定的 JSON 格式輸出。
    --- 使用者訊息 ---
    「{message_text}」
    --- 輔助資訊 ---
    今天日期：「{json.dumps(str(app.config.get('TODAYS_DATE', '')))}」 (YYYY-MM-DD 格式)
    --- 輸出格式 (請嚴格遵守此 JSON 結構) ---
    {{
      "project_name": "(從訊息中提取的專案名稱，若無則為 '未指定')",
      "task_name": "(從訊息中提取的任務名稱)",
      "assignee": "(從訊息中提取的負責人，若無則為 '未指定')",
      "due_date": "(從訊息中提取的預計完成日期，並轉換為 YYYY-MM-DD 格式，若無則為 'NULL')"
    }}
    --- 指示 ---
    1.  請務必回傳一個完整的 JSON 物件。
    2.  `task_name` 必須提取，如果訊息不清楚，請提取最關鍵的動詞片語。
    3.  日期解析：請將「明天」、「下週五」、「11/20」等自然語言轉換為 YYYY-MM-DD 格式。
    4.  `due_date` 如果真的無法解析，請回傳 `NULL` (字串)。
    5.  不要在 JSON 以外添加任何 'json' 標記或任何說明文字。
    """
    try:
        response = gemini_model.generate_content(prompt)
        json_string = response.text.strip().replace("```json", "").replace("```", "")
        task_data = json.loads(json_string)
        app.logger.info(f"Gemini 解析結果: {task_data}")
    except Exception as e:
        app.logger.error(f"Gemini API 或 JSON 解析失敗: {e}")
        app.logger.error(f"Gemini 原始回傳: {response.text if 'response' in locals() else 'N/A'}")
        raise Exception(f"Gemini 解析資料失敗，請檢查 Prompt 或 API 金鑰。")
    try:
        new_row = [
            task_data.get('project_name', '未指定'), # A 專案名稱
            task_data.get('task_name', '未提取'),    # B 任務名稱
            task_data.get('assignee', '未指定'),     # C 負責人
            "待辦",                                # D 任務狀態 (預設)
            "中",                                 # E 優先級 (預設)
            app.config.get('TODAYS_DATE', ''),      # F 開始日期 (預定義)
            task_data.get('due_date', None),      # G 預計完成日期
            None,                                 # H 實際完成日期 (留空)
            0,                                    # I 進度% (預設 0)
            None,                                 # J 問題與風險 (留空)
            "完成時",                             # K 回報頻率 (預設)
            message_text                          # L 任務描述 / 備註 (原始指令)
        ]
        worksheet.append_row(new_row)
        app.logger.info("成功寫入 Google Sheet")
    except Exception as e:
        app.logger.error(f"Google Sheet 寫入失敗: {e}")
        raise Exception(f"寫入 Google Sheet 失敗，請檢查 service_account.json 權限。")
    reply_text = f"""
    ✅ 任務已新增：
    
    專案：{new_row[0]}
    任務：{new_row[1]}
    負責人：{new_row[2]}
    狀態：{new_row[3]}
    截止日期：{new_row[6] if new_row[6] else '未定'}
    """
    return reply_text.strip()


# --- 5. 核心功能：被動查詢 (Handle Query Task) ---
# (此函式與 v3 完全相同，故折疊)
def handle_query_task(message_text):
    app.logger.info(f"執行任務查詢：{message_text}")
    try:
        all_data = worksheet.get_all_records()
        if not all_data:
            return "資料庫中尚無任何任務可供查詢。"
        df = pd.DataFrame(all_data)
        data_context = df.to_markdown(index=False)
    except Exception as e:
        app.logger.error(f"讀取 Google Sheet 失敗: {e}")
        raise Exception("讀取 Google Sheet 資料時發生錯誤。")
    prompt = f"""
    你是一位專業的專案管理助理。
    
    以下是目前 Google Sheets 中的所有任務資料（使用 Markdown 格式）：
    --- 資料開始 ---
    {data_context}
    --- 資料結束 ---
    
    輔助資訊：
    - 今天的日期是：{app.config.get('TODAYS_DATE')}
    
    請根據上述資料，用繁體中文、條列式的方式，簡潔地回答以下使用者的問題。
    
    --- 使用者的問題 ---
    「{message_text}」
    
    --- 你的回答 ---
    """
    try:
        response = gemini_model.generate_content(prompt)
        app.logger.info("Gemini 查詢回答成功。")
        return response.text.strip()
    except Exception as e:
        app.logger.error(f"Gemini 查詢 API 失敗: {e}")
        raise Exception("Gemini 在分析資料時發生錯誤。")


# --- 6. 【v4 新功能】 核心功能：主動推播 (Send Daily Summary) ---

def send_daily_summary():
    """
    產生並推播每日晨報：
    1. 檢查 USER_ID_TO_PUSH 是否存在。
    2. 讀取 GSheet 資料。
    3. 呼叫 Gemini 產生晨報。
    4. 使用 Push API 傳送晨報。
    """
    app.logger.info(f"開始執行每日晨報推播任務...")
    
    # 1. 檢查 USER_ID
    if not USER_ID_TO_PUSH:
        app.logger.warning("USER_ID_TO_PUSH 未設定，無法推播晨報。")
        return

    # 2. 讀取 GSheet 資料
    try:
        all_data = worksheet.get_all_records()
        if not all_data:
            app.logger.info("資料庫中無任務，不推播晨報。")
            return
        df = pd.DataFrame(all_data)
        # 篩選掉「已完成」的任務，只看相關的
        active_tasks_df = df[df['任務狀態'] != '已完成']
        if active_tasks_df.empty:
            app.logger.info("無進行中任務，不推播晨報。")
            return
        
        data_context = active_tasks_df.to_markdown(index=False)
    except Exception as e:
        app.logger.error(f"推播任務：讀取 GSheet 失敗: {e}")
        return # 函式失敗

    # 3. 呼叫 Gemini 產生晨報
    prompt = f"""
    你是一位專業且語氣親切的專案管理助理。
    
    以下是目前 Google Sheets 中所有「未完成」的任務資料：
    --- 資料開始 ---
    {data_context}
    --- 資料結束 ---
    
    輔助資訊：
    - 今天的日期是：{app.config.get('TODAYS_DATE')}
    
    請根據上述資料，產生一份「今日晨報 (Daily Summary)」。
    晨報內容應包含：
    1.  一句早安問候。
    2.  【今日到期】: 列出「預計完成日期」等於今天的任務。
    3.  【即將到期】: 列出未來7天內到期的任務。
    4.  【進行中任務】: 列出其他「進行中」或「待辦」的任務。
    5.  如果沒有上述任務，就說「今天沒有緊急任務」。
    6.  最後加上一句鼓勵的話。
    
    請使用繁體中文，格式清晰易讀。
    --- 你的晨報 ---
    """
    
    try:
        response = gemini_model.generate_content(prompt)
        summary_text = response.text.strip()
    except Exception as e:
        app.logger.error(f"推播任務：Gemini 產生晨報失敗: {e}")
        summary_text = f"❌ 產生晨報失敗：{e}" # 即使失敗也要推播錯誤訊息

    # 4. 使用 Push API 傳送晨報
    try:
        global_line_bot_api.push_message_with_http_info(
            PushMessageRequest(
                to=USER_ID_TO_PUSH,
                messages=[TextMessage(text=summary_text)]
            )
        )
        app.logger.info(f"成功推播晨報給 {USER_ID_TO_PUSH}")
    except Exception as e:
        app.logger.error(f"推播任務：Push API 失敗: {e}")


# --- 7. 啟動伺服器與排程器 ---

if __name__ == "__main__":
    # 【v4 新增】 啟動排程器
    scheduler = BackgroundScheduler(timezone='Asia/Taipei') # 設定時區
    
    # 新增您指定的任務：平日 (mon-fri) 早上 8:00
    scheduler.add_job(
        send_daily_summary, 
        trigger='cron', 
        day_of_week='mon-fri', 
        hour=8, 
        minute=0
    )
    
    # 啟動排程器
    scheduler.start()
    app.logger.info("排程器已啟動... (mon-fri, 8:00am, Asia/Taipei)")
    
    # 啟動 Flask 伺服器
    # 我們關閉 debug=True 並允許 '0.0.0.0'，這是為未來部署到雲端做準備
    # debug=True 會導致排程器運行兩次，所以我們在 v4 中將其關閉
    app.run(host='0.0.0.0', port=5001, debug=False)
    '''
import os
import json
from datetime import datetime
from dotenv import load_dotenv, set_key
from flask import Flask, request, abort

# --- LINE Bot SDK ---
from linebot.v3 import (
    WebhookHandler
)
from linebot.v3.exceptions import (
    InvalidSignatureError
)
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    PushMessageRequest,
    TextMessage
)
from linebot.v3.webhooks import (
    MessageEvent,
    TextMessageContent
)

# --- Google SDKs ---
import google.generativeai as genai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# --- Scheduler ---
from apscheduler.schedulers.background import BackgroundScheduler

# --- 1. 初始設定 ---

load_dotenv()
dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
app = Flask(__name__)

# --- LINE 設定 ---
LINE_CHANNEL_SECRET = os.getenv('LINE_CHANNEL_SECRET')
LINE_CHANNEL_ACCESS_TOKEN = os.getenv('LINE_CHANNEL_ACCESS_TOKEN')
line_configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)
global_api_client = ApiClient(line_configuration)
global_line_bot_api = MessagingApi(global_api_client)
USER_ID_TO_PUSH = os.getenv('USER_ID_TO_PUSH')

# --- Gemini 設定 ---
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')
genai.configure(api_key=GEMINI_API_KEY)
# 我們使用您帳戶可用的模型
gemini_model = genai.GenerativeModel('gemini-2.5-flash') 

# --- Google Sheets 設定 ---
GOOGLE_SHEET_NAME = os.getenv('GOOGLE_SHEET_NAME')
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
client = gspread.authorize(creds)
try:
    spreadsheet = client.open(GOOGLE_SHEET_NAME)
    worksheet = spreadsheet.sheet1
    app.logger.info(f"成功開啟 Google Sheet: {GOOGLE_SHEET_NAME}")
    # 【v5 優化】 啟動時讀取一次欄位名稱，供 AI 篩選器使用
    app.config['SHEET_HEADERS'] = worksheet.row_values(1)
    app.logger.info(f"成功讀取欄位: {app.config['SHEET_HEADERS']}")
except gspread.exceptions.SpreadsheetNotFound:
    app.logger.error(f"錯誤：找不到名稱為 '{GOOGLE_SHEET_NAME}' 的 Google Sheet")
    exit()

app.config['TODAYS_DATE'] = datetime.now().strftime('%Y-%m-%d')


# --- 2. 建立 Webhook 監聽路徑 (與 v4 相同) ---

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        app.logger.error("簽名錯誤")
        abort(400)
    return 'OK'

# --- 3. 處理「文字訊息」事件 (與 v4 相同) ---

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    global USER_ID_TO_PUSH
    message_text = event.message.text
    reply_token = event.reply_token
    user_id = event.source.user_id
    reply_text = ""

    if not USER_ID_TO_PUSH:
        USER_ID_TO_PUSH = user_id
        set_key(dotenv_path, "USER_ID_TO_PUSH", USER_ID_TO_PUSH)
        app.logger.info(f"成功儲存 User ID: {user_id} 到 .env 檔案")
        reply_text = f"你好！我已經將您的 User ID 設為推播目標。\n\n"
    elif USER_ID_TO_PUSH != user_id:
        app.logger.warning(f"偵測到不同的 User ID: {user_id} (已設定: {USER_ID_TO_PUSH})")

    # 【指令路由器 (Router)】
    if "提醒" in message_text or "新增" in message_text or "指派" in message_text or "待辦" in message_text:
        try:
            reply_text += handle_write_task(message_text)
        except Exception as e:
            app.logger.error(f"智能寫入失敗: {e}")
            reply_text += f"❌ 任務新增失敗：\n{e}"
    
    elif "查詢" in message_text or "進度" in message_text or "有什麼" in message_text or "幫我找" in message_text:
        try:
            # 【v5 優化】 呼叫新的二階段查詢
            reply_text += handle_query_task_v5(message_text)
        except Exception as e:
            app.logger.error(f"任務查詢失敗: {e}")
            reply_text += f"❌ 任務查詢失敗：\n{e}"

    elif "晨報" in message_text or "summary" in message_text:
        try:
            app.logger.info(f"手動觸發晨報，推播給: {USER_ID_TO_PUSH}")
            send_daily_summary()
            reply_text += "✅ 晨報已發送，請檢查您的 LINE 訊息。"
        except Exception as e:
            app.logger.error(f"手動觸發晨報失敗: {e}")
            reply_text += f"❌ 晨報發送失敗：\n{e}"
    
    else:
        reply_text += "您好，我是一個專案管理助理。\n\n- 新增任務：「提醒我...」\n- 查詢任務：「查詢...」\n- 測試晨報：「晨報」"

    # 統一回覆
    global_line_bot_api.reply_message_with_http_info(
        ReplyMessageRequest(
            reply_token=reply_token,
            messages=[TextMessage(text=reply_text.strip())]
        )
    )

# --- 4. 核心功能：智能寫入 (Handle Write Task) ---
# (此函式與 v4 完全相同，故折疊)
def handle_write_task(message_text):
    app.logger.info(f"執行智能寫入：{message_text}")
    prompt = f"""
    你是一個專案管理 AI 助理。請從以下使用者訊息中，解析並提取資訊，並嚴格按照指定的 JSON 格式輸出。
    --- 使用者訊息 ---
    「{message_text}」
    --- 輔助資訊 ---
    今天日期：「{json.dumps(str(app.config.get('TODAYS_DATE', '')))}」 (YYYY-MM-DD 格式)
    --- 輸出格式 (請嚴格遵守此 JSON 結構) ---
    {{
      "project_name": "(從訊息中提取的專案名稱，若無則為 '未指定')",
      "task_name": "(從訊息中提取的任務名稱)",
      "assignee": "(從訊息中提取的負責人，若無則為 '未指定')",
      "due_date": "(從訊息中提取的預計完成日期，並轉換為 YYYY-MM-DD 格式，若無則為 'NULL')"
    }}
    --- 指示 ---
    1.  請務必回傳一個完整的 JSON 物件。
    2.  `task_name` 必須提取，如果訊息不清楚，請提取最關鍵的動詞片語。
    3.  日期解析：請將「明天」、「下週五」、「11/20」等自然語言轉換為 YYYY-MM-DD 格式。
    4.  `due_date` 如果真的無法解析，請回傳 `NULL` (字串)。
    5.  不要在 JSON 以外添加任何 'json' 標記或任何說明文字。
    """
    try:
        response = gemini_model.generate_content(prompt)
        json_string = response.text.strip().replace("```json", "").replace("```", "")
        task_data = json.loads(json_string)
        app.logger.info(f"Gemini 解析結果: {task_data}")
    except Exception as e:
        app.logger.error(f"Gemini API 或 JSON 解析失敗: {e}")
        app.logger.error(f"Gemini 原始回傳: {response.text if 'response' in locals() else 'N/A'}")
        raise Exception(f"Gemini 解析資料失敗，請檢查 Prompt 或 API 金鑰。")
    try:
        new_row = [
            task_data.get('project_name', '未指定'), # A 專案名稱
            task_data.get('task_name', '未提取'),    # B 任務名稱
            task_data.get('assignee', '未指定'),     # C 負責人
            "待辦",                                # D 任務狀態 (預設)
            "中",                                 # E 優先級 (預設)
            app.config.get('TODAYS_DATE', ''),      # F 開始日期 (預定義)
            task_data.get('due_date', None),      # G 預計完成日期
            None,                                 # H 實際完成日期 (留空)
            0,                                    # I 進度% (預設 0)
            None,                                 # J 問題與風險 (留空)
            "完成時",                             # K 回報頻率 (預設)
            message_text                          # L 任務描述 / 備註 (原始指令)
        ]
        worksheet.append_row(new_row)
        app.logger.info("成功寫入 Google Sheet")
    except Exception as e:
        app.logger.error(f"Google Sheet 寫入失敗: {e}")
        raise Exception(f"寫入 Google Sheet 失敗，請檢查 service_account.json 權限。")
    reply_text = f"""
    ✅ 任務已新增：
    
    專案：{new_row[0]}
    任務：{new_row[1]}
    負責人：{new_row[2]}
    狀態：{new_row[3]}
    截止日期：{new_row[6] if new_row[6] else '未定'}
    """
    return reply_text.strip()


# --- 5. 【v5 核心優化】 二階段智能查詢 ---

# 輔助函式：第一階段 AI (快速解析)
def parse_query_to_filters(message_text):
    """
    【AI 呼叫 1】
    使用者的自然語言 -> 結構化 JSON 篩選條件
    """
    app.logger.info(f"V5 查詢 - 階段 1: 解析篩選條件...")
    
    # 從 app config 讀取欄位名稱
    sheet_columns = app.config['SHEET_HEADERS']
    
    prompt = f"""
    你是一個查詢解析器。請將使用者的「自然語言查詢」轉換為一個 JSON 物件，用於篩選 pandas DataFrame。
    
    --- 可用的欄位 (Columns) ---
    {json.dumps(sheet_columns)}
    
    --- 使用者的查詢 ---
    「{message_text}」
    
    --- 你的 JSON 輸出 ---
    請只回傳 JSON 物件，key 必須是上述可用的欄位之一，value 則是對應的篩選值。
    如果使用者沒有指定某個欄位的篩選，請不要包含該 key。
    
    範例 1:
    查詢: "查詢『醫療AI專案』的進度"
    輸出: {{"專案名稱": "醫療AI專案"}}
    
    範例 2:
    查詢: "幫我找所有『張經理』負責的任務"
    輸出: {{"負責人": "張經理"}}
    
    範例 3:
    查詢: "有哪些任務卡住了？"
    輸出: {{"任務狀態": "卡住"}}
    
    範例 4:
    查詢: "所有『醫療AI專案』中『已完成』的任務"
    輸出: {{"專案名稱": "醫療AI專案", "任務狀態": "已完成"}}
    """
    
    try:
        response = gemini_model.generate_content(prompt)
        json_string = response.text.strip().replace("```json", "").replace("```", "")
        filters = json.loads(json_string)
        app.logger.info(f"V5 查詢 - 階段 1 成功: {filters}")
        return filters
    except Exception as e:
        app.logger.error(f"V5 查詢 - 階段 1 失敗: {e}")
        # 如果解析失敗，我們就退回 V4 做法 (空篩選)
        return {}

# 輔助函式：第二階段 AI (快速總結)
def summarize_filtered_data(filtered_df, message_text):
    """
    【AI 呼叫 2】
    (Python 篩選過的) 少量資料 -> AI 總結報告
    """
    app.logger.info(f"V5 查詢 - 階段 3: 總結 {len(filtered_df)} 筆資料...")
    
    data_context = filtered_df.to_markdown(index=False)
    
    prompt = f"""
    你是一位專業的專案管理助理。
    
    以下是使用者**篩選後**的任務資料（使用 Markdown 格式）：
    --- 資料開始 ---
    {data_context}
    --- 資料結束 ---
    
    輔助資訊：
    - 今天的日期是：{app.config.get('TODAYS_DATE')}
    
    請根據上述資料，用繁體中文、條列式的方式，簡潔地回答以下使用者的原始問題。
    
    --- 使用者的原始問題 ---
    「{message_text}」
    
    --- 你的回答 ---
    """
    try:
        response = gemini_model.generate_content(prompt)
        app.logger.info("V5 查詢 - 階段 3 成功。")
        return response.text.strip()
    except Exception as e:
        app.logger.error(f"V5 查詢 - 階段 3 失敗: {e}")
        raise Exception("Gemini 在總結資料時發生錯誤。")


# 主要查詢函式 (v5)
def handle_query_task_v5(message_text):
    """
    處理被動查詢任務 (V5 優化版)：
    1. 【AI 1】 將自然語言轉為 JSON 篩選條件
    2. 【GSheet】 讀取所有資料
    3. 【Python】 使用 JSON 條件在本地篩選資料
    4. 【AI 2】 將篩選後的少量資料交給 AI 總結
    """
    app.logger.info(f"執行 V5 任務查詢：{message_text}")
    
    # 1. 【AI 1】 解析篩選條件
    filters = parse_query_to_filters(message_text)
    
    # 2. 【GSheet】 讀取所有資料
    try:
        all_data = worksheet.get_all_records()
        if not all_data:
            return "資料庫中尚無任何任務可供查詢。"
        df = pd.DataFrame(all_data)
    except Exception as e:
        app.logger.error(f"讀取 Google Sheet 失敗: {e}")
        raise Exception("讀取 Google Sheet 資料時發生錯誤。")

    # 3. 【Python】 本地篩選
    if not filters:
        # 如果 AI 1 解析失敗，退回 V4 做法 (總結全部)
        app.logger.warning("V5 查詢 - 階段 1 解析失敗，退回 V4 總結全部資料。")
        filtered_df = df
    else:
        app.logger.info(f"V5 查詢 - 階段 2: 執行本地篩選 {filters}")
        # 建立一個布林遮罩 (boolean mask)
        mask = pd.Series(True, index=df.index)
        for key, value in filters.items():
            if key in df.columns:
                # 確保比較時型別一致 (例如 GSheet 讀入的都是字串)
                mask &= (df[key].astype(str) == str(value))
        filtered_df = df[mask]
    
    if filtered_df.empty:
        return f"根據您的篩選條件，找不到任何相關任務。\n(篩選條件: {filters})"

    # 4. 【AI 2】 總結
    return summarize_filtered_data(filtered_df, message_text)


# --- 6. 核心功能：主動推播 (Send Daily Summary) ---
# (此函式與 v4 完全相同，故折疊)
def send_daily_summary():
    app.logger.info(f"開始執行每日晨報推播任務...")
    if not USER_ID_TO_PUSH:
        app.logger.warning("USER_ID_TO_PUSH 未設定，無法推播晨報。")
        return
    try:
        all_data = worksheet.get_all_records()
        if not all_data:
            app.logger.info("資料庫中無任務，不推播晨報。")
            return
        df = pd.DataFrame(all_data)
        active_tasks_df = df[df['任務狀態'] != '已完成']
        if active_tasks_df.empty:
            app.logger.info("無進行中任務，不推播晨報。")
            return
        data_context = active_tasks_df.to_markdown(index=False)
    except Exception as e:
        app.logger.error(f"推播任務：讀取 GSheet 失敗: {e}")
        return
    prompt = f"""
    你是一位專業且語氣親切的專案管理助理。
    以下是目前 Google Sheets 中所有「未完成」的任務資料：
    --- 資料開始 ---
    {data_context}
    --- 資料結束 ---
    輔助資訊：
    - 今天的日期是：{app.config.get('TODAYS_DATE')}
    請根據上述資料，產生一份「今日晨報 (Daily Summary)」。
    晨報內容應包含：
    1.  一句早安問候。
    2.  【今日到期】: 列出「預計完成日期」等於今天的任務。
    3.  【即將到期】: 列出未來7天內到期的任務。
    4.  【進行中任務】: 列出其他「進行中」或「待辦」的任務。
    5.  如果沒有上述任務，就說「今天沒有緊急任務」。
    6.  最後加上一句鼓勵的話。
    請使用繁體中文，格式清晰易讀。
    --- 你的晨報 ---
    """
    try:
        response = gemini_model.generate_content(prompt)
        summary_text = response.text.strip()
    except Exception as e:
        app.logger.error(f"推播任務：Gemini 產生晨報失敗: {e}")
        summary_text = f"❌ 產生晨報失敗：{e}"
    try:
        global_line_bot_api.push_message_with_http_info(
            PushMessageRequest(
                to=USER_ID_TO_PUSH,
                messages=[TextMessage(text=summary_text)]
            )
        )
        app.logger.info(f"成功推播晨報給 {USER_ID_TO_PUSH}")
    except Exception as e:
        app.logger.error(f"推播任務：Push API 失敗: {e}")


# --- 7. 啟動伺服器與排程器 ---
if __name__ == "__main__":
    scheduler = BackgroundScheduler(timezone='Asia/Taipei')
    scheduler.add_job(
        send_daily_summary, 
        trigger='cron', 
        day_of_week='mon-fri', 
        hour=7, 
        minute=30,
    )
    scheduler.start()
    app.logger.info("排程器已啟動... (mon-fri, 7:30am, Asia/Taipei)")
    app.run(host='0.0.0.0', port=5001, debug=False)