import streamlit as st
import pandas as pd
import gspread
from streamlit_gsheets import GSheetsConnection
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import time

# ================= 設定區 =================
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1gpq9Cye25rmPgyOt508L1sBvlIpPis45R09vn0uy434/edit"

SYSTEM_COLS = ['Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']
REQUIRED_COLS = ['ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', '反思會', '反思表', '家長/監護人']

st.set_page_config(page_title="雲端實習津貼系統", layout="wide", page_icon="☁️")

# ================= 連線設定 =================
@st.cache_resource
def get_write_client():
    try:
        creds_dict = dict(st.secrets["connections"]["gsheets"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        client = gspread.service_account_from_dict(creds_dict)
        return client
    except Exception as e:
        st.error(f"連線設定錯誤: {e}")
        st.stop()

conn = st.connection("gsheets", type=GSheetsConnection)

# ================= 核心函式 =================
def fetch_data_from_cloud(sheet_name):
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl=0)
        if not df.empty:
            df.columns = df.columns.str.strip()
            if 'ID序號' in df.columns:
                df['ID序號'] = df['ID序號'].astype(str)
            else:
                if len(df.columns) > 0:
                    df.rename(columns={df.columns[0]: 'ID序號'}, inplace=True)
                    df['ID序號'] = df['ID序號'].astype(str)

            for col in SYSTEM_COLS:
                if col not in df.columns: df[col] = ''
            df = df.fillna('')
        else:
            df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)
        return df
    except Exception as e:
        return pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

# ================= 主程式開始 =================
st.title("☁️ 實習津貼管理系統 (V34 上傳優化版)")

if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    
    st.divider()
    
    try:
        gc = get_write_client()
        sh = gc.open_by_url(SPREADSHEET_URL)
        sheet_names = [ws.title for ws in sh.worksheets()]
        selected_sheet_name = st.selectbox("📂 選擇工作表 (資料來源)", sheet_names, index=0)
    except Exception as e:
        st.error(f"連線失敗: {e}")
        st.stop()

    need_refresh = st.button("🔄 重新整理資料")
    
    if need_refresh or st.session_state.df_main is None or st.session_state.current_sheet != selected_sheet_name:
        with st.spinner(f"正在讀取「{selected_sheet_name}」..."):
            st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
            st.session_state.current_sheet = selected_sheet_name
            if need_refresh: st.success("資料已更新！")

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名才能開始操作。")
    st.stop()

df = st.session_state.df_main

try:
    worksheet = sh.worksheet(selected_sheet_name)
except:
    st.warning("工作表讀取中...")
    st.stop()

# ================= 分頁功能 =================
tab_upload, tab_prepare, tab_confirm, tab_manage = st.tabs([
    "📥 建立新工作表", 
    "📄 [1] 準備匯出", 
    "✅ [2] 確認領取", 
    "🛠️ 資料管理 (刪除)"
])

# -------------------------------------------
# TAB 1: 上載新資料 (修復重點)
# -------------------------------------------
with tab_upload:
    st.subheader("📥 上傳 Excel 並建立獨立工作表")
    
    # 1. 檔案上傳器 (設定 key 以便清除)
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=['xlsx', 'xls'], key="uploader_key")
    
    # 2. 文字輸入框 (設定 key 以便清除)
    new_sheet_name = st.text_input("請輸入新工作表名稱 (例如: 2024_第一期)", placeholder="請輸入名稱...", key="new_sheet_input")
    
    if uploaded_file and new_sheet_name:
        # 檢查名稱是否重複
        if new_sheet_name in sheet_names:
            st.error(f"⚠️ 工作表名稱「{new_sheet_name}」已存在！請更換名稱。")
        else:
            try:
                new_df = pd.read_excel(uploaded_file)
                if len(new_df.columns) >= 9:
                    mapping = {
                        new_df.columns[0]: 'ID序號', new_df.columns[1]: '編號',
                        new_df.columns[2]: '姓名(中文)', new_df.columns[3]: '姓名(英文)',
                        new_df.columns[4]: '電話', new_df.columns[5]: '實習日數',
                        new_df.columns[6]: '反思會', new_df.columns[7]: '反思表',
                        new_df.columns[8]: '家長/監護人'
                    }
                    new_df.rename(columns=mapping, inplace=True)
                    valid_cols = [c for c in REQUIRED_COLS if c in new_df.columns]
                    new_df = new_df[valid_cols]
                    
                    for col in SYSTEM_COLS: new_df[col] = ''
