import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import io
import time

# ================= 設定區 =================
# 請再次確認您的 Google Sheet 網址
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/您的_GOOGLE_SHEET_ID/edit"

# 系統與必要欄位
SYSTEM_COLS = ['Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']
REQUIRED_COLS = ['ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', '反思會', '反思表', '家長/監護人']

st.set_page_config(page_title="雲端實習津貼系統", layout="wide", page_icon="☁️")

# ================= 核心修復：建立寫入專用的連線 =================
@st.cache_resource
def get_write_client():
    """
    建立一個原生的 gspread 客戶端，專門用於「寫入」和「精確修改」。
    它會直接讀取 secrets.toml 中的設定。
    """
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    # 從 Streamlit secrets 讀取憑證
    # 注意：這裡對應 secrets.toml 中的 [connections.gsheets]
    try:
        creds_dict = dict(st.secrets["connections"]["gsheets"])
        
        # 清理並準備憑證字典 (移除不必要的鍵值以免報錯)
        # ServiceAccountCredentials 需要標準的 JSON 結構
        clean_creds = {
            "type": creds_dict.get("type"),
            "project_id": creds_dict.get("project_id"),
            "private_key_id": creds_dict.get("private_key_id"),
            "private_key": creds_dict.get("private_key"),
            "client_email": creds_dict.get("client_email"),
            "client_id": creds_dict.get("client_id"),
            "auth_uri": creds_dict.get("auth_uri"),
            "token_uri": creds_dict.get("token_uri"),
            "auth_provider_x509_cert_url": creds_dict.get("auth_provider_x509_cert_url"),
            "client_x509_cert_url": creds_dict.get("client_x509_cert_url"),
        }
        
        creds = ServiceAccountCredentials.from_json_keyfile_dict(clean_creds, scope)
        client = gspread.authorize(creds)
        return client
    except KeyError:
        st.error("❌ 找不到 Secrets 設定，請檢查 secrets.toml 是否有 [connections.gsheets] 區塊")
        st.stop()
    except Exception as e:
        st.error(f"❌ 憑證授權失敗: {e}")
        st.stop()

# ================= 建立讀取連線 (快速讀取用) =================
conn = st.connection("gsheets", type=GSheetsConnection)

# ================= 主程式開始 =================
st.title("☁️ 實習津貼管理系統 (雲端修復版)")

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    
    st.divider()
    
    # 1. 取得工作表列表 (使用寫入專用的 client)
    try:
        gc = get_write_client()
        sh = gc.open_by_url(SPREADSHEET_URL)
        sheet_names = [ws.title for ws in sh.worksheets()]
        selected_sheet_name = st.selectbox("📂 選擇工作表", sheet_names)
    except Exception as e:
        st.error(f"無法連線至 Google Sheets。\n錯誤: {e}")
        st.stop()

    if st.button("🔄 重新整理資料"):
        st.cache_data.clear()
        st.rerun()

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名才能開始操作。")
    st.stop()

# --- 讀取資料 (使用 conn 快速讀取) ---
try:
    # ttl=0 確保不快取，每次抓最新
    df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=selected_sheet_name, ttl=0)
    
    # 資料清理
    if not df.empty:
        df['ID序號'] = df['ID序號'].astype(str)
        for col in SYSTEM_COLS:
            if col not in df.columns: df[col] = ''
        df = df.fillna('')
    else:
        df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

except Exception as e:
    st.error(f"讀取資料失敗: {e}")
    st.stop()

# 取得寫入用的 worksheet 物件
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
                new_df = new_df[REQUIRED_COLS]
                
                for col in SYSTEM_COLS: new_df[col] = ''
                new_df['ID序號'] = new_df['ID序號'].astype(str)
                new_df = new_df.fillna('') # 確保沒有 NaN
                
                st.write("預覽:", new_df.head())
                
                if st.button("🚀 確認上傳"):
                    with st.spinner("寫入雲端中..."):
                        worksheet.append_rows(new_df.values.tolist())
                        st.success(f"成功新增 {len(new_df)} 筆資料！")
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
            
            # 取得 Header Index
            header = worksheet.row_values(1)
            try:
                col_doc_idx = header.index('DocGeneratedDate') + 1
                col_staff_idx = header.index('ResponsibleStaff') + 1
            except:
                st.error("雲端表格缺少系統欄位 (DocGeneratedDate/ResponsibleStaff)")
                st.stop()

            progress_bar = st.progress(0)
            status_text = st.empty()
            export_list = []
            
            for i, (idx, row) in enumerate(selected.iterrows()):
                target_id = row['ID序號']
                try:
                    # 使用 gspread 的 find 進行精確定位
                    cell = worksheet.find(target_id, in_column=1)
                    if cell:
                        worksheet.update_cell(cell.row, col_doc_idx, today)
                        worksheet.update_cell(cell.row, col_staff_idx, staff_name)
                        
                        rec = row.to_dict()
                        del rec['選取']
                        rec['StaffName'] = staff_name
                        rec['TodayDate'] = today
                        export_list.append(rec)
                        status_text.text(f"已更新: {row['姓名(中文)']}")
                except Exception as e:
                    st.warning(f"更新 ID {target_id} 失敗: {e}")
                
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
                time.sleep(2)
                st.rerun()

# -------------------------------------------
# TAB 3: 確認領取
# -------------------------------------------
with tab_confirm:
    st.subheader("✅ 步驟二：確認領取")
    
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
                    st.error("缺少 Collected 或 CollectedDate 欄位")
                    st.stop()
                
                prog = st.progress(0)
                status = st.empty()
                for i, (idx, row) in enumerate(selected.iterrows()):
                    try:
                        cell = worksheet.find(row['ID序號'], in_column=1)
                        if cell:
                            worksheet.update_cell(cell.row, col_col_idx, 'Y')
                            worksheet.update_cell(cell.row, col_date_idx, now_str)
                            status.text(f"已確認: {row['姓名(中文)']}")
                    except: pass
                    prog.progress((i + 1) / len(selected))
                
                st.success("更新完成！")
                time.sleep(1)
                st.rerun()
    
    with col2:
        if st.button("↩️ 退回至準備匯出"):
            selected = edited_confirm[edited_confirm["確認"] == True]
            if not selected.empty:
                if st.checkbox("確定要退回嗎？(這會清除文件日期)"):
                    header = worksheet.row_values(1)
                    col_doc_idx = header.index('DocGeneratedDate') + 1
                    col_staff_idx = header.index('ResponsibleStaff') + 1
                    for idx, row in selected.iterrows():
                        try:
                            cell = worksheet.find(row['ID序號'], in_column=1)
                            if cell:
                                worksheet.update_cell(cell.row, col_doc_idx, "")
                                worksheet.update_cell(cell.row, col_staff_idx, "")
                        except: pass
                    st.success("已退回")
                    time.sleep(1)
                    st.rerun()

# -------------------------------------------
# TAB 4: 總覽
# -------------------------------------------
with tab_history:
    st.subheader("📜 資料總覽")
    st.dataframe(df)
