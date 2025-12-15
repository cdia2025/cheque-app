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
                    new_df['ID序號'] = new_df['ID序號'].astype(str)
                    new_df = new_df.fillna('') 
                    
                    st.write("預覽:", new_df.head())
                    
                    if st.button("🚀 建立新表並上傳"):
                        with st.spinner("正在建立新工作表..."):
                            # 建立並寫入
                            new_ws = sh.add_worksheet(title=new_sheet_name, rows=len(new_df)+50, cols=20)
                            data_to_write = [new_df.columns.tolist()] + new_df.values.tolist()
                            new_ws.update('A1', data_to_write)
                            
                            st.success(f"成功建立「{new_sheet_name}」！")
                            
                            # --- 關鍵修正：清空輸入框狀態 ---
                            # 這樣下次重新整理時，輸入框會變回空白，就不會觸發「重複名稱」警告
                            st.session_state["new_sheet_input"] = "" 
                            # 如果想連檔案也清空，可以 uncomment 下面這行
                            # st.session_state["uploader_key"] = None 
                            
                            st.info("系統將於 2 秒後重新整理...")
                            time.sleep(2)
                            
                            # 清除快取並重整
                            st.cache_data.clear()
                            st.rerun()
                else:
                    st.error("欄位不足 9 欄")
            except Exception as e:
                st.error(f"錯誤: {e}")

# -------------------------------------------
# TAB 2: 準備匯出
# -------------------------------------------
with tab_prepare:
    st.subheader(f"📄 準備匯出 ({selected_sheet_name})")
    
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
                header = worksheet.row_values(1)
                try:
                    col_doc_idx = header.index('DocGeneratedDate') + 1
                    col_staff_idx = header.index('ResponsibleStaff') + 1
                except:
                    st.error("雲端表格缺少系統欄位")
                    st.stop()

                progress_bar = st.progress(0)
                export_list = []
                
                for i, (idx, row) in enumerate(selected.iterrows()):
                    target_id = row['ID序號']
                    try:
                        cell = worksheet.find(target_id, in_column=1)
                        if cell:
                            worksheet.update_cell(cell.row, col_doc_idx, today)
                            worksheet.update_cell(cell.row, col_staff_idx, staff_name)
                            
                            org_idx = df[df['ID序號'] == target_id].index
                            if not org_idx.empty:
                                st.session_state.df_main.loc[org_idx, 'DocGeneratedDate'] = today
                                st.session_state.df_main.loc[org_idx, 'ResponsibleStaff'] = staff_name

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
                    st.download_button(label="📥 下載 MailMerge_Source.xlsx", data=buffer.getvalue(), file_name="MailMerge_Source.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.info("介面已同步更新。")

# -------------------------------------------
# TAB 3: 確認領取
# -------------------------------------------
with tab_confirm:
    st.subheader(f"✅ 確認領取 ({selected_sheet_name})")
    
    if 'Collected' in df.columns:
        mask_confirm = ((df['DocGeneratedDate'] != '') & (df['Collected'] != 'Y'))
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
                    col_col_idx = header.index('Collected') + 1
                    col_date_idx = header.index('CollectedDate') + 1
                    
                    prog = st.progress(0)
                    for i, (idx, row) in enumerate(selected.iterrows()):
                        try:
                            cell = worksheet.find(row['ID序號'], in_column=1)
                            if cell:
                                worksheet.update_cell(cell.row, col_col_idx, 'Y')
                                worksheet.update_cell(cell.row, col_date_idx, now_str)
                                org_idx = df[df['ID序號'] == row['ID序號']].index
                                st.session_state.df_main.loc[org_idx, 'Collected'] = 'Y'
                                st.session_state.df_main.loc[org_idx, 'CollectedDate'] = now_str
                        except: pass
                        prog.progress((i + 1) / len(selected))
                    st.success("更新完成！")
                    st.rerun()
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
                            except: pass
                        st.success("已退回")
                        st.rerun()

# -------------------------------------------
# TAB 4: 資料管理 (刪除功能)
# -------------------------------------------
with tab_manage:
    st.subheader(f"🛠️ 資料管理與刪除 ({selected_sheet_name})")
    st.warning("⚠️ 請小心操作，此處的刪除動作將「永久」從 Google Sheets 移除資料！")
    
    df_manage = df.copy()
    df_manage.insert(0, "刪除", False)
    
    edited_manage = st.data_editor(
        df_manage,
        column_config={"刪除": st.column_config.CheckboxColumn(required=True, label="選取刪除")},
        hide_index=True,
        key="editor_manage"
    )
    
    col_del_sel, col_del_all = st.columns(2)
    
    with col_del_sel:
        if st.button("🗑️ 刪除「已勾選」的資料", type="secondary"):
            selected_del = edited_manage[edited_manage["刪除"] == True]
            if selected_del.empty:
                st.warning("請先勾選要刪除的資料")
            else:
                count = len(selected_del)
                if st.checkbox(f"確定要刪除這 {count} 筆資料嗎？(無法復原)", key="confirm_del_sel"):
                    with st.spinner("正在刪除..."):
                        rows_to_delete = []
                        for idx, row in selected_del.iterrows():
                            try:
                                cell = worksheet.find(row['ID序號'], in_column=1)
                                if cell:
                                    rows_to_delete.append(cell.row)
                            except: pass
                        
                        rows_to_delete.sort(reverse=True)
                        for r_idx in rows_to_delete:
                            worksheet.delete_rows(r_idx)
                            
                        st.success(f"已刪除 {len(rows_to_delete)} 筆資料。")
                        time.sleep(1)
                        st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
                        st.rerun()

    with col_del_all:
        if st.button("💥 清空本工作表所有資料", type="primary"):
            st.warning("這將刪除此工作表內的所有內容（保留第一列標題）。")
            if st.checkbox("我確定要清空整張表 (TYPE CONFIRM)", key="confirm_clear"):
                with st.spinner("正在清空..."):
                    headers = worksheet.row_values(1)
                    worksheet.clear()
                    worksheet.update('A1', [headers])
                    st.success("工作表已清空！")
                    time.sleep(1)
                    st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
                    st.rerun()
