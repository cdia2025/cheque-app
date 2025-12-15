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
with tab_upload:
    st.subheader("📥 上傳 Excel 並附加到目前工作表")
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            new_df = pd.read_excel(uploaded_file)
            if len(new_df.columns) >= 9:
                # 欄位對應
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
                if 'ID序號' in new_df.columns:
                    new_df['ID序號'] = new_df['ID序號'].astype(str)
                new_df = new_df.fillna('') 
                
                st.write("預覽:", new_df.head())
                
                if st.button("🚀 確認上傳"):
                    with st.spinner("寫入雲端中..."):
                        worksheet.append_rows(new_df.values.tolist())
                        st.success(f"成功新增 {len(new_df)} 筆資料！正在重新整理...")
                        # 強制重讀
                        st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
                        time.sleep(1)
                        st.rerun()
            else:
                st.error("欄位不足 9 欄")
        except Exception as e:
            st.error(f"錯誤: {e}")

# -------------------------------------------
# TAB 2: 準備匯出
# -------------------------------------------
with tab_prepare:
    st.subheader("📄 步驟一：匯出 Mail Merge 資料")
    
    # 確保欄位存在
    if '反思會' in df.columns:
        mask_ready = (
            (df['反思會'].astype(str).str.upper() == 'Y') & 
            (df['反思表'].astype(str).str.upper() == 'Y') & 
            (df['DocGeneratedDate'] == '')
        )
        df_ready = df[mask_ready].copy()
        
        df_ready.insert(0, "選取", False)
        edited_ready = st.data_editor(
            df_ready,
            column_config={"選取": st.column_config.CheckboxColumn(required=True)},
            disabled=[c for c in df.columns if c != "選取"],
            hide_index=True,
            key="editor_ready"
        )
        
        if st.button("📤 匯出 & 更新狀態", type="primary"):
            selected = edited_ready[edited_ready["選取"] == True]
            if selected.empty:
                st.warning("未選取")
            else:
                today = datetime.now().strftime("%Y-%m-%d")
                
                # 取得欄位 Index
                header = worksheet.row_values(1)
                try:
                    col_doc_idx = header.index('DocGeneratedDate') + 1
                    col_staff_idx = header.index('ResponsibleStaff') + 1
                except:
                    st.error("雲端表格缺少 DocGeneratedDate 欄位")
                    st.stop()

                progress_bar = st.progress(0)
                export_list = []
                
                # 批次更新邏輯
                for i, (idx, row) in enumerate(selected.iterrows()):
                    target_id = row['ID序號']
                    try:
                        cell = worksheet.find(target_id, in_column=1)
                        if cell:
                            # 1. 寫入雲端
                            worksheet.update_cell(cell.row, col_doc_idx, today)
                            worksheet.update_cell(cell.row, col_staff_idx, staff_name)
                            
                            # 2. 重要：同步更新本地 Session State (避免為了顯示結果又去讀 Google)
                            # 找出原始 df 中的 index
                            org_idx = df[df['ID序號'] == target_id].index
                            if not org_idx.empty:
                                st.session_state.df_main.loc[org_idx, 'DocGeneratedDate'] = today
                                st.session_state.df_main.loc[org_idx, 'ResponsibleStaff'] = staff_name

                            # 準備下載資料
                            rec = row.to_dict()
                            del rec['選取']
                            rec['StaffName'] = staff_name
                            rec['TodayDate'] = today
                            export_list.append(rec)
                    except: pass
                    progress_bar.progress((i + 1) / len(selected))
                
                if export_list:
                    out_df = pd.DataFrame(export_list)
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        out_df.to_excel(writer, index=False)
                    
                    st.success(f"完成！已更新 {len(export_list)} 筆。")
                    st.download_button(
                        label="📥 下載 MailMerge_Source.xlsx",
                        data=buffer.getvalue(),
                        file_name="MailMerge_Source.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    # 不用 sleep 和 rerun，因為我們已經手動更新了 session state，畫面下次互動會自動變
                    st.info("介面資料已同步更新。")

# -------------------------------------------
# TAB 3: 確認領取
# -------------------------------------------
with tab_confirm:
    st.subheader("✅ 步驟二：確認領取")
    
    if 'Collected' in df.columns:
        mask_confirm = (
            (df['DocGeneratedDate'] != '') & 
            (df['Collected'] != 'Y')
        )
        df_confirm = df[mask_confirm].copy()
        
        df_confirm.insert(0, "確認", False)
        edited_confirm = st.data_editor(
            df_confirm,
            column_config={"確認": st.column_config.CheckboxColumn(required=True)},
            disabled=[c for c in df.columns if c != "確認"],
            hide_index=True,
            key="editor_confirm"
        )
        
        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("✅ 確認已取票", type="primary"):
                selected = edited_confirm[edited_confirm["確認"] == True]
                if selected.empty:
                    st.warning("未選取")
                else:
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    header = worksheet.row_values(1)
                    try:
                        col_col_idx = header.index('Collected') + 1
                        col_date_idx = header.index('CollectedDate') + 1
                    except:
                        st.error("缺少 Collected 欄位")
                        st.stop()
                    
                    prog = st.progress(0)
                    for i, (idx, row) in enumerate(selected.iterrows()):
                        try:
                            cell = worksheet.find(row['ID序號'], in_column=1)
                            if cell:
                                worksheet.update_cell(cell.row, col_col_idx, 'Y')
                                worksheet.update_cell(cell.row, col_date_idx, now_str)
                                
                                # 同步更新 Session State
                                org_idx = df[df['ID序號'] == row['ID序號']].index
                                st.session_state.df_main.loc[org_idx, 'Collected'] = 'Y'
                                st.session_state.df_main.loc[org_idx, 'CollectedDate'] = now_str
                        except: pass
                        prog.progress((i + 1) / len(selected))
                    
                    st.success("更新完成！")
                    st.rerun() # 這裡需要 rerun 來刷新列表
        
        with col2:
            if st.button("↩️ 退回至準備匯出"):
                selected = edited_confirm[edited_confirm["確認"] == True]
                if not selected.empty:
                    if st.checkbox("確定要退回嗎？(清除日期)"):
                        header = worksheet.row_values(1)
                        col_doc_idx = header.index('DocGeneratedDate') + 1
                        col_staff_idx = header.index('ResponsibleStaff') + 1
                        for idx, row in selected.iterrows():
                            try:
                                cell = worksheet.find(row['ID序號'], in_column=1)
                                if cell:
                                    worksheet.update_cell(cell.row, col_doc_idx, "")
                                    worksheet.update_cell(cell.row, col_staff_idx, "")
                                    
                                    # 更新 Session
                                    org_idx = df[df['ID序號'] == row['ID序號']].index
                                    st.session_state.df_main.loc[org_idx, 'DocGeneratedDate'] = ''
                                    st.session_state.df_main.loc[org_idx, 'ResponsibleStaff'] = ''
                            except: pass
                        st.success("已退回")
                        st.rerun()

# -------------------------------------------
# TAB 4: 總覽
# -------------------------------------------
with tab_history:
    st.subheader("📜 資料總覽")
    st.dataframe(df)
