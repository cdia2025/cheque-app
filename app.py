import streamlit as st
import pandas as pd
import gspread
from streamlit_gsheets import GSheetsConnection
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import time

# ================= 設定區 =================
# 您的 Google Sheet 網址
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1gpq9Cye25rmPgyOt508L1sBvlIpPis45R09vn0uy434/edit"

# 系統與必要欄位
SYSTEM_COLS = ['Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']
REQUIRED_COLS = ['ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', '反思會', '反思表', '家長/監護人']

st.set_page_config(page_title="雲端實習津貼系統", layout="wide", page_icon="☁️")

# ================= 連線設定 =================
@st.cache_resource
def get_write_client():
    """建立寫入專用的 gspread 客戶端"""
    try:
        creds_dict = dict(st.secrets["connections"]["gsheets"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
        # 使用 gspread 新版驗證方法
        client = gspread.service_account_from_dict(creds_dict)
        return client
    except Exception as e:
        st.error(f"連線設定錯誤: {e}")
        st.stop()

# 讀取連線 (保留但改用手動觸發)
conn = st.connection("gsheets", type=GSheetsConnection)

# ================= 核心函式：讀取資料並存入 Session =================
def fetch_data_from_cloud(sheet_name):
    """從 Google Sheet 讀取資料，並處理格式"""
    try:
        # 使用 ttl=0 強制讀取最新，但這個函式我們只會在必要時呼叫
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl=0)
        
        if not df.empty:
            df.columns = df.columns.str.strip() # 去除欄位空白
            
            # 欄位對應與修正
            if 'ID序號' in df.columns:
                df['ID序號'] = df['ID序號'].astype(str)
            else:
                # 若找不到 ID，自動抓第一欄
                df.rename(columns={df.columns[0]: 'ID序號'}, inplace=True)
                df['ID序號'] = df['ID序號'].astype(str)

            for col in SYSTEM_COLS:
                if col not in df.columns: df[col] = ''
            df = df.fillna('')
        else:
            df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)
            
        return df
    except Exception as e:
        st.error(f"讀取失敗 (Quota Exceeded?): {e}")
        return pd.DataFrame()

# ================= 主程式開始 =================
st.title("☁️ 實習津貼管理系統 (V32 防流量限制版)")

# --- 初始化 Session State ---
# 這是避免 429 錯誤的關鍵：資料存在這裡，不會一直去煩 Google
if 'df_main' not in st.session_state:
    st.session_state.df_main = None
if 'current_sheet' not in st.session_state:
    st.session_state.current_sheet = None

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    
    st.divider()
    
    # 1. 取得工作表列表 (這個動作消耗很少 quota，可以保留)
    try:
        gc = get_write_client()
        sh = gc.open_by_url(SPREADSHEET_URL)
        sheet_names = [ws.title for ws in sh.worksheets()]
        selected_sheet_name = st.selectbox("📂 選擇工作表", sheet_names)
    except Exception as e:
        st.error(f"連線失敗: {e}")
        st.stop()

    # 2. 讀取/重整按鈕
    # 邏輯：如果換了工作表，或者按了重整，才去讀取 Google
    need_refresh = st.button("🔄 重新整理資料 (從雲端讀取)")
    
    if need_refresh or st.session_state.df_main is None or st.session_state.current_sheet != selected_sheet_name:
        with st.spinner("正在從 Google 下載資料..."):
            st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
            st.session_state.current_sheet = selected_sheet_name
            # 如果是按按鈕觸發的，顯示成功訊息
            if need_refresh:
                st.success("資料已更新！")

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名才能開始操作。")
    st.stop()

# 使用 Session 中的資料
df = st.session_state.df_main

# 取得寫入用的 worksheet 物件 (只建立連線物件，不讀取資料，不耗 Quota)
try:
    worksheet = sh.worksheet(selected_sheet_name)
except:
    st.error("無法取得工作表物件")
    st.stop()

# ================= 分頁功能 =================
tab_upload, tab_prepare, tab_confirm, tab_history = st.tabs([
    "📥 上載新資料", 
    "📄 [1] 準備匯出 (Mail Merge)", 
    "✅ [2] 確認領取", 
    "📜 資料總覽"
])

# -------------------------------------------
# TAB 1: 上載新資料
# -------------------------------------------
with tab_uploa
