import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import time

# ================= 設定區 =================
# 1. 請填入您的 Google Sheet ID
SHEET_ID = "1gpq9Cye25rmPgyOt508L1sBvlIpPis45R09vn0uy434/edit?gid=0#gid=0" 

# 2. 金鑰檔案名稱
JSON_KEYFILE = "secrets.json"

# 3. 系統欄位 (程式會自動管理這些欄位)
SYSTEM_COLS = ['Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']

# 4. 必要的基礎欄位 (上傳的 Excel 必須包含這些)
REQUIRED_COLS = ['ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', '反思會', '反思表', '家長/監護人']

st.set_page_config(page_title="雲端實習津貼系統", layout="wide", page_icon="☁️")

# ================= 連線函式 =================
@st.cache_resource
def get_google_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # 支援 Streamlit Cloud Secrets 或 本地 JSON
    if "gcp_service_account" in st.secrets:
        creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEYFILE, scope)
    return gspread.authorize(creds)

def get_worksheet(sheet_name):
    client = get_google_client()
    sh = client.open_by_key(SHEET_ID)
    return sh.worksheet(sheet_name), sh

# ================= 介面開始 =================
st.title("☁️ 實習津貼管理系統 (Google Sheets 連動版)")

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    
    # 1. 職員登入
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    
    st.divider()
    
    # 2. 選擇工作表
    try:
        client = get_google_client()
        sh = client.open_by_key(SHEET_ID)
        sheet_names = [ws.title for ws in sh.worksheets()]
        selected_sheet = st.selectbox("📂 選擇工作表", sheet_names)
    except Exception as e:
        st.error("無法連線至 Google Sheets，請檢查 ID 與權限。")
        st.stop()

    if st.button("🔄 重新整理資料"):
        st.cache_data.clear()
        st.rerun()

# 檢查職員姓名
if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名才能開始操作。")
    st.stop()

# --- 讀取資料 ---
try:
    worksheet = sh.worksheet(selected_sheet)
    data = worksheet.get_all_records()
    # 轉為 DataFrame，並確保 ID 是字串
    df = pd.DataFrame(data)
    if not df.empty:
        df['ID序號'] = df['ID序號'].astype(str)
        # 確保系統欄位存在 (若 GSheet 漏了，這裡補上空值以免報錯)
        for col in SYSTEM_COLS:
            if col not in df.columns:
                df[col] = ''
    else:
        # 如果是空表，建立空 DataFrame
        df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

except Exception as e:
    st.error(f"讀取資料失敗: {e}")
    st.stop()

# ================= 分頁功能 =================
tab_upload, tab_prepare, tab_confirm, tab_history = st.tabs([
    "📥 上載新資料", 
    "📄 [1] 準備匯出 (Mail Merge)", 
    "✅ [2] 確認領取", 
    "📜 資料總覽"
])

# -------------------------------------------
# TAB 1: 上載新資料 (附加到 Google Sheet)
# -------------------------------------------
with tab_upload:
    st.subheader("📥 上傳新的 Excel 名單")
    st.info("上傳的資料將會「附加 (Append)」到目前 Google Sheet 的最下方。")
    
    uploaded_file = st.file_uploader("請選擇 Excel 檔案", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            new_df = pd.read_excel(uploaded_file)
            
            # 1. 欄位檢查與對應
            # 這裡假設使用者上傳的檔案欄位順序是固定的 (Index 0-8)
            # 如果欄位名稱不同，強制改名以符合系統標準
            if len(new_df.columns) >= 9:
                mapping = {
                    new_df.columns[0]: 'ID序號',
                    new_df.columns[1]: '編號',
                    new_df.columns[2]: '姓名(中文)',
                    new_df.columns[3]: '姓名(英文)',
                    new_df.columns[4]: '電話',
                    new_df.columns[5]: '實習日數',
                    new_df.columns[6]: '反思會',
                    new_df.columns[7]: '反思表',
                    new_df.columns[8]: '家長/監護人'
                }
                new_df.rename(columns=mapping, inplace=True)
                
                # 只保留需要的欄位
                new_df = new_df[REQUIRED_COLS]
                
                # 補上系統欄位 (空值)
                for col in SYSTEM_COLS:
                    new_df[col] = ''
                
                # 確保 ID 是字串
                new_df['ID序號'] = new_df['ID序號'].astype(str)
                
                # 預覽
                st.write("預覽即將上傳的資料 (前 5 筆):")
                st.dataframe(new_df.head())
                
                if st.button("🚀 確認上傳並寫入 Google Sheets", type="primary"):
                    with st.spinner("正在寫入雲端..."):
                        # 將 DataFrame 轉為 List of Lists
                        values = new_df.values.tolist()
                        # 使用 append_rows 一次寫入，效率高
                        worksheet.append_rows(values)
                        st.success(f"成功新增 {len(values)} 筆資料！")
                        time.sleep(1)
                        st.rerun() # 重新整理
            else:
                st.error("上傳的檔案欄位不足 9 欄，請檢查格式。")
                
        except Exception as e:
            st.error(f"處理檔案時發生錯誤: {e}")

# -------------------------------------------
# TAB 2: 準備匯出 (產生 Mail Merge Source)
# -------------------------------------------
with tab_prepare:
    st.subheader("📄 步驟一：篩選並匯出 Mail Merge 資料")
    
    if df.empty:
        st.warning("目前沒有資料。")
    else:
        # 篩選：符合資格 (Y/Y) 且 尚未生成文件
        mask_ready = (
            (df['反思會'].astype(str).str.upper() == 'Y') & 
            (df['反思表'].astype(str).str.upper() == 'Y') & 
            (df['DocGeneratedDate'] == '')
        )
        df_ready = df[mask_ready].copy()
        
        # 讓使用者勾選
        df_ready.insert(0, "選取", False)
        edited_ready = st.data_editor(
            df_ready,
            column_config={"選取": st.column_config.CheckboxColumn(required=True)},
            disabled=[c for c in df.columns if c != "選取"],
            hide_index=True,
            key="editor_ready"
        )
        
        if st.button("📤 匯出資料 & 更新狀態", type="primary"):
            selected = edited_ready[edited_ready["選取"] == True]
            
            if selected.empty:
                st.warning("請至少勾選一人")
            else:
                today = datetime.now().strftime("%Y-%m-%d")
                
                # 準備更新 GSheet
                # 為了效率，我們這裡使用 cell.value 查找 (若資料量大建議優化)
                # 這裡假設 ID 在第 1 欄 (col 1)
                
                # 取得 header index
                header = worksheet.row_values(1)
                try:
                    col_doc_idx = header.index('DocGeneratedDate') + 1
                    col_staff_idx = header.index('ResponsibleStaff') + 1
                except:
                    st.error("Google Sheet 缺少系統欄位，請檢查標題列。")
                    st.stop()

                progress_bar = st.progress(0)
                export_list = []
                total = len(selected)
                
                for i, (idx, row) in enumerate(selected.iterrows()):
                    target_id = row['ID序號']
                    
                    # 尋找列數 (使用 find)
                    try:
                        cell = worksheet.find(target_id, in_column=1)
                        if cell:
                            # 更新雲端
                            worksheet.update_cell(cell.row, col_doc_idx, today)
                            worksheet.update_cell(cell.row, col_staff_idx, staff_name)
                            
                            # 準備匯出資料
                            record = row.to_dict()
                            del record['選取']
                            record['StaffName'] = staff_name
                            record['TodayDate'] = today
                            export_list.append(record)
                    except Exception as e:
                        st.error(f"更新 ID {target_id} 時發生錯誤: {e}")
                    
                    progress_bar.progress((i + 1) / total)
                
                if export_list:
                    # 產生 Excel 下載
                    export_df = pd.DataFrame(export_list)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        export_df.to_excel(writer, index=False)
                    
                    st.success(f"已更新 {len(export_list)} 筆資料！狀態改為「待領取」。")
                    st.download_button(
                        label="📥 下載 Mail Merge 專用檔 (MailMerge_Source.xlsx)",
                        data=output.getvalue(),
                        file_name="MailMerge_Source.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    time.sleep(2)
                    st.rerun()

# -------------------------------------------
# TAB 3: 確認領取
# -------------------------------------------
with tab_confirm:
    st.subheader("✅ 步驟二：確認領取支票")
    
    if df.empty:
        st.warning("無資料")
    else:
        # 篩選：已生成文件 但 未領取
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
                    st.warning("請選擇人員")
                else:
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    header = worksheet.row_values(1)
                    col_col_idx = header.index('Collected') + 1
                    col_date_idx = header.index('CollectedDate') + 1
                    
                    prog = st.progress(0)
                    total = len(selected)
                    
                    for i, (idx, row) in enumerate(selected.iterrows()):
                        cell = worksheet.find(row['ID序號'], in_column=1)
                        if cell:
                            worksheet.update_cell(cell.row, col_col_idx, 'Y')
                            worksheet.update_cell(cell.row, col_date_idx, now_str)
                        prog.progress((i + 1) / total)
                    
                    st.success("更新完成！")
                    time.sleep(1)
                    st.rerun()
        
        with col2:
            # 退回功能 (Undo)
            if st.button("↩️ 退回至準備匯出 (清除日期)"):
                selected = edited_confirm[edited_confirm["確認"] == True]
                if selected.empty:
                    st.warning("請選擇要退回的人員")
                else:
                    if st.checkbox("確定要退回嗎？這會清除該人員的文件日期。", value=False):
                        header = worksheet.row_values(1)
                        col_doc_idx = header.index('DocGeneratedDate') + 1
                        col_staff_idx = header.index('ResponsibleStaff') + 1
                        
                        for i, (idx, row) in enumerate(selected.iterrows()):
                            cell = worksheet.find(row['ID序號'], in_column=1)
                            if cell:
                                worksheet.update_cell(cell.row, col_doc_idx, "")
                                worksheet.update_cell(cell.row, col_staff_idx, "")
                        st.success("已退回至「步驟一」。")
                        time.sleep(1)
                        st.rerun()

# -------------------------------------------
# TAB 4: 資料總覽
# -------------------------------------------
with tab_history:
    st.subheader("📜 目前母檔總覽")
    st.dataframe(df)