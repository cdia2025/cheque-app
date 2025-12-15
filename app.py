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

# ================= 核心：萬能 ID 清洗函式 (V44 新增) =================
def clean_id(val):
    """
    將各種奇形怪狀的 ID (數字、浮點數字串、含空白字串) 統一轉為乾淨的文字。
    範例: 
    101     -> "101"
    " 101 " -> "101"
    101.0   -> "101"
    "101.0" -> "101"
    """
    if val is None: return ""
    s = str(val).strip()
    if s == "": return ""
    # 處理 Excel 常見的 .0 結尾
    if s.endswith(".0"):
        return s[:-2]
    return s

# ================= 連線設定 =================
@st.cache_resource
def get_write_client():
    try:
        creds_dict = dict(st.secrets["connections"]["gsheets"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        return gspread.service_account_from_dict(creds_dict)
    except Exception as e:
        st.error(f"連線設定錯誤: {e}")
        st.stop()

conn = st.connection("gsheets", type=GSheetsConnection)

# ================= 核心函式 (快取) =================
@st.cache_data(ttl=600)
def get_sheet_names_cached():
    try:
        gc = get_write_client()
        sh = gc.open_by_url(SPREADSHEET_URL)
        return [ws.title for ws in sh.worksheets()]
    except: return []

@st.cache_data(ttl=600)
def fetch_data_cached(sheet_name):
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name)
        if not df.empty:
            df.columns = df.columns.str.strip()
            
            # 處理欄位名稱
            if 'ID序號' not in df.columns and len(df.columns) > 0:
                df.rename(columns={df.columns[0]: 'ID序號'}, inplace=True)
            
            # --- V44 修正：使用 clean_id 清洗 ID ---
            if 'ID序號' in df.columns:
                df['ID序號'] = df['ID序號'].apply(clean_id)
            
            for col in SYSTEM_COLS:
                if col not in df.columns: df[col] = ''
            df = df.fillna('')
        else:
            df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)
        return df
    except:
        return pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

# ================= 統計計算 =================
def calculate_stats(df):
    if df.empty or '反思會' not in df.columns:
        return {'總人數': 0, '待匯出': 0, '待領取': 0, '已完成': 0, '不符資格': 0}
    c1 = df['反思會'].astype(str).str.strip().str.upper()
    c2 = df['反思表'].astype(str).str.strip().str.upper()
    doc = df['DocGeneratedDate'].astype(str).str.strip()
    done = df['Collected'].astype(str).str.strip().str.upper()
    is_eligible = (c1 == 'Y') & (c2 == 'Y')
    return {
        '總人數': len(df),
        '待匯出': (is_eligible & (doc == '')).sum(),
        '待領取': ((doc != '') & (done != 'Y')).sum(),
        '已完成': (done == 'Y').sum(),
        '不符資格': ((~is_eligible) & (doc == '')).sum()
    }

# ================= 主程式 =================
st.title("☁️ 實習津貼管理系統 (V44 格式統整版)")

# Session State
if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    st.divider()
    
    sheet_names = get_sheet_names_cached()
    if not sheet_names:
        st.error("讀取失敗，請檢查連線。")
        st.stop()

    if 'last_selected_sheet' not in st.session_state:
        st.session_state.last_selected_sheet = sheet_names[0]
        
    idx = 0
    if st.session_state.last_selected_sheet in sheet_names:
        idx = sheet_names.index(st.session_state.last_selected_sheet)
        
    selected_sheet_name = st.selectbox("📂 選擇工作表", sheet_names, index=idx)
    st.session_state.last_selected_sheet = selected_sheet_name

    if st.button("🔄 重新整理資料"):
        st.cache_data.clear()
        st.rerun()

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名。")
    st.stop()

df = fetch_data_cached(selected_sheet_name)

# ================= 分頁 =================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📥 建立新表", "📄 [1] 匯出", "✅ [2] 領取", 
    "🚫 [3] 不符", "🛠️ 管理", "✏️ 修改", "📊 統計"
])

# ---------------- Tab 1: 建立新表 ----------------
with tab1:
    st.subheader("📥 上傳 Excel")
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=['xlsx', 'xls'], key="upl")
    new_sheet_name = st.text_input("輸入新工作表名稱", placeholder="2024_第一期", key="new_s_in")
    if st.button("🚀 建立並上傳", type="primary"):
        if uploaded_file and new_sheet_name:
            if new_sheet_name in sheet_names: st.error("名稱已存在！")
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
                        new_df = new_df[REQUIRED_COLS]
                        for col in SYSTEM_COLS: new_df[col] = ''
                        
                        # V44 修正：上傳時也清洗 ID
                        new_df['ID序號'] = new_df['ID序號'].apply(clean_id)
                        new_df = new_df.fillna('')
                        
                        with st.spinner("建立中..."):
                            gc = get_write_client()
                            sh = gc.open_by_url(SPREADSHEET_URL)
                            new_ws = sh.add_worksheet(title=new_sheet_name, rows=len(new_df)+50, cols=20)
                            # 為了確保 Google Sheet 也是乾淨的文字格式，轉成 str 再上傳
                            data_to_write = [new_df.columns.tolist()] + new_df.astype(str).values.tolist()
                            new_ws.update(data_to_write)
                            
                            st.success(f"成功建立「{new_sheet_name}」！")
                            time.sleep(2); st.cache_data.clear(); st.rerun()
                    else: st.error("欄位不足")
                except Exception as e: st.error(f"錯誤: {e}")
        else: st.error("請填寫名稱並選擇檔案")

# ---------------- Tab 2: 準備匯出 ----------------
with tab
