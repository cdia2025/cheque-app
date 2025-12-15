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

# ================= 連線設定 (寫入用) =================
@st.cache_resource
def get_write_client():
    """建立寫入專用的 gspread 客戶端"""
    try:
        creds_dict = dict(st.secrets["connections"]["gsheets"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        return gspread.service_account_from_dict(creds_dict)
    except Exception as e:
        st.error(f"連線設定錯誤: {e}")
        st.stop()

# 讀取連線
conn = st.connection("gsheets", type=GSheetsConnection)

# ================= 核心函式 (加入快取機制以解決 429 錯誤) =================

@st.cache_data(ttl=600)  # 快取 10 分鐘，除非手動重整
def get_sheet_names_cached():
    """取得工作表清單 (快取版)"""
    try:
        gc = get_write_client()
        sh = gc.open_by_url(SPREADSHEET_URL)
        return [ws.title for ws in sh.worksheets()]
    except Exception as e:
        return []

@st.cache_data(ttl=600) # 快取 10 分鐘，除非手動重整
def fetch_data_cached(sheet_name):
    """讀取工作表內容 (快取版)"""
    try:
        # 使用 conn.read 會自動處理 Pandas 轉換
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name)
        
        if not df.empty:
            df.columns = df.columns.str.strip()
            # 處理 ID
            if 'ID序號' in df.columns:
                df['ID序號'] = df['ID序號'].astype(str).str.strip()
            else:
                if len(df.columns) > 0:
                    df.rename(columns={df.columns[0]: 'ID序號'}, inplace=True)
                    df['ID序號'] = df['ID序號'].astype(str).str.strip()
            
            # 補齊欄位
            for col in SYSTEM_COLS:
                if col not in df.columns: df[col] = ''
            
            df = df.fillna('')
        else:
            df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)
        return df
    except Exception as e:
        # 發生錯誤回傳空表，避免程式崩潰
        return pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

# ================= 主程式 =================
st.title("☁️ 實習津貼管理系統 (V39 流量優化版)")

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    st.divider()
    
    # 1. 取得工作表清單 (使用快取，不消耗 Quota)
    sheet_names = get_sheet_names_cached()
    
    if not sheet_names:
        st.error("無法讀取工作表，請稍後再試或檢查連線。")
        st.stop()

    # 選擇工作表
    if 'last_selected_sheet' not in st.session_state:
        st.session_state.last_selected_sheet = sheet_names[0]
        
    # 如果列表變更了，重置選擇
    idx = 0
    if st.session_state.last_selected_sheet in sheet_names:
        idx = sheet_names.index(st.session_state.last_selected_sheet)
        
    selected_sheet_name = st.selectbox("📂 選擇工作表", sheet_names, index=idx)
    st.session_state.last_selected_sheet = selected_sheet_name

    # 2. 重新整理按鈕 (這是唯一清除快取的地方)
    if st.button("🔄 重新整理資料 (Clear Cache)"):
        st.cache_data.clear() # 清除所有快取
        st.rerun()

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名。")
    st.stop()

# --- 讀取資料 (使用快取) ---
# 這裡不會每次都連線 Google，除非快取過期或按了重整
df = fetch_data_cached(selected_sheet_name)

# ================= 分頁 =================
tab_upload, tab_prepare, tab_confirm, tab_manage = st.tabs([
    "📥 建立新表", "📄 [1] 匯出", "✅ [2] 領取", "🛠️ 管理"
])

# ---------------- Tab 1: 建立新表 ----------------
with tab_upload:
    st.subheader("📥 上傳 Excel 並建立獨立工作表")
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=['xlsx', 'xls'], key="upl")
    new_sheet_name = st.text_input("輸入新工作表名稱", placeholder="2024_第一期", key="new_s_in")
    
    if st.button("🚀 建立並上傳", type="primary"):
        if uploaded_file and new_sheet_name:
            if new_sheet_name in sheet_names:
                st.error("名稱已存在！")
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
                        new_df['ID序號'] = new_df['ID序號'].astype(str)
                        new_df = new_df.fillna('')
                        
                        with st.spinner("建立中..."):
                            # 寫入時才建立連線
                            gc = get_write_client()
                            sh = gc.open_by_url(SPREADSHEET_URL)
                            new_ws = sh.add_worksheet(title=new_sheet_name, rows=len(new_df)+50, cols=20)
                            new_ws.update([new_df.columns.tolist()] + new_df.values.tolist())
                            
                            st.success(f"成功建立「{new_sheet_name}」！")
                            time.sleep(2) # 等待 Google 同步
                            st.cache_data.clear() # 清除快取以顯示新表
                            st.rerun()
                    else: st.error("欄位不足")
                except Exception as e: st.error(f"錯誤: {e}")
        else: st.error("請填寫名稱並選擇檔案")

# ---------------- Tab 2: 準備匯出 ----------------
with tab_prepare:
    st.subheader(f"📄 準備匯出 ({selected_sheet_name})")
    if '反思會' in df.columns:
        mask_ready = ((df['反思會'].astype(str).str.upper() == 'Y') & (df['反思表'].astype(str).str.upper() == 'Y') & (df['DocGeneratedDate'] == ''))
        df_ready = df[mask_ready].copy()
        df_ready.insert(0, "選取", False)
        edited_ready = st.data_editor(df_ready, column_config={"選取": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df.columns if c!="選取"], hide_index=True, key="ed_ready")
        
        if st.button("📤 匯出 & 更新狀態", type="primary"):
            sel = edited_ready[edited_ready["選取"]==True]
            if not sel.empty:
                try:
                    # 寫入操作：建立連線
                    gc = get_write_client()
                    worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                    
                    today = datetime.now().strftime("%Y-%m-%d")
                    head = worksheet.row_values(1)
                    c_doc = head.index('DocGeneratedDate')+1
                    c_staff = head.index('ResponsibleStaff')+1
                    
                    # 批次取得 Cloud IDs (減少 API 呼叫)
                    cloud_ids = [str(x).strip() for x in worksheet.col_values(1)]
                    
                    prog = st.progress(0)
                    ex_list = []
                    
                    for i, (idx, row) in enumerate(sel.iterrows()):
                        target_id = str(row['ID序號']).strip()
                        if target_id in cloud_ids:
                            # 找出所有符合的 row index (加 1 因為 list 從 0 開始，sheet 從 1 開始)
                            # 這裡只取第一個匹配的
                            row_num = cloud_ids.index(target_id) + 1
                            worksheet.update_cell(row_num, c_doc, today)
                            worksheet.update_cell(row_num, c_staff, staff_name)
                            
                            rec = row.to_dict(); del rec['選取']; rec.update({'StaffName':staff_name, 'TodayDate':today})
                            ex_list.append(rec)
                        
                        prog.progress((i+1)/len(sel))
                    
                    if ex_list:
                        out = io.BytesIO()
                        pd.DataFrame(ex_list).to_excel(out, index=False)
                        st.download_button("📥 下載 MailMerge Source", out.getvalue(), "MailMerge_Source.xlsx")
                        st.success("完成！請按重新整理以查看更新。")
                        time.sleep(1)
                        st.cache_data.clear() # 寫入後清除快取
                except Exception as e: 
                    st.error(f"寫入錯誤 (可能流量過大，請稍後再試): {e}")

# ---------------- Tab 3: 確認領取 ----------------
with tab_confirm:
    st.subheader(f"✅ 確認領取 ({selected_sheet_name})")
    if 'Collected' in df.columns:
        mask_conf = ((df['DocGeneratedDate']!='') & (df['Collected']!='Y'))
        df_conf = df[mask_conf].copy()
        df_conf.insert(0, "確認", False)
        ed_conf = st.data_editor(df_conf, column_config={"確認": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df.columns if c!="確認"], hide_index=True, key="ed_conf")
        
        if st.button("✅ 確認已取票", type="primary"):
            sel = ed_conf[ed_conf["確認"]==True]
            if not sel.empty:
                try:
                    gc = get_write_client()
                    worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                    
                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    head = worksheet.row_values(1)
                    c_col = head.index('Collected')+1
                    c_date = head.index('CollectedDate')+1
                    
                    cloud_ids = [str(x).strip() for x in worksheet.col_values(1)]
                    
                    prog = st.progress(0)
                    for i, (idx, row) in enumerate(sel.iterrows()):
                        target_id = str(row['ID序號']).strip()
                        if target_id in cloud_ids:
                            row_num = cloud_ids.index(target_id) + 1
                            worksheet.update_cell(row_num, c_col, 'Y')
                            worksheet.update_cell(row_num, c_date, now)
                        prog.progress((i+1)/len(sel))
                    st.success("更新完成！")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"錯誤: {e}")

# ---------------- Tab 4: 刪除管理 (精簡版) ----------------
with tab_manage:
    st.subheader(f"🛠️ 工作表管理 - {selected_sheet_name}")
    st.info("已移除單筆刪除功能，僅保留刪除整張工作表。")
    
    st.divider()
    
    if st.button("🔥 請求刪除本工作表"):
        if len(sheet_names) <= 1:
            st.error("Google Sheets 至少需保留一張表，無法刪除。")
        else:
            if 'confirm_del_sheet' not in st.session_state:
                st.session_state.confirm_del_sheet = True
    
    if st.session_state.get('confirm_del_sheet', False):
        st.warning(f"確定要永久刪除「{selected_sheet_name}」分頁嗎？")
        col_yes, col_no = st.columns(2)
        with col_yes:
            if st.button("🔴 是，確認刪除", key="btn_confirm_sheet"):
                with st.spinner("刪除中..."):
                    gc = get_write_client()
                    sh = gc.open_by_url(SPREADSHEET_URL)
                    worksheet = sh.worksheet(selected_sheet_name)
                    
                    sh.del_worksheet(worksheet)
                    st.success("工作表已刪除")
                    st.session_state.confirm_del_sheet = False
                    time.sleep(2)
                    st.cache_data.clear()
                    st.rerun()
        with col_no:
            if st.button("取消", key="btn_cancel_sheet"):
                st.session_state.confirm_del_sheet = False
                st.rerun()
