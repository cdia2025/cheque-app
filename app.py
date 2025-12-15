import streamlit as st
import pandas as pd
import gspread
from streamlit_gsheets import GSheetsConnection
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import io
import time

# ================= 設定區 =================
# 請確認這裡填的是 Google Sheet 的網址
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
        
        client = gspread.service_account_from_dict(creds_dict)
        return client
    except Exception as e:
        st.error(f"連線設定錯誤: {e}")
        st.stop()

# 讀取連線
conn = st.connection("gsheets", type=GSheetsConnection)

# ================= 核心函式 =================
def fetch_data_from_cloud(sheet_name):
    """從 Google Sheet 讀取資料"""
    try:
        # ttl=0 強制讀取最新
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl=0)
        
        if not df.empty:
            df.columns = df.columns.str.strip()
            
            # 處理 ID 欄位
            if 'ID序號' in df.columns:
                df['ID序號'] = df['ID序號'].astype(str)
            else:
                if len(df.columns) > 0:
                    df.rename(columns={df.columns[0]: 'ID序號'}, inplace=True)
                    df['ID序號'] = df['ID序號'].astype(str)

            # 補齊系統欄位
            for col in SYSTEM_COLS:
                if col not in df.columns: df[col] = ''
            
            df = df.fillna('')
        else:
            df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)
        return df
    except Exception as e:
        # 若發生錯誤回傳空 DataFrame
        return pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

# ================= 主程式開始 =================
st.title("☁️ 實習津貼管理系統 (V35 資料刪除增強版)")

# 初始化 Session State
if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None
if 'new_sheet_input' not in st.session_state: st.session_state.new_sheet_input = ""

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    
    st.divider()
    
    try:
        gc = get_write_client()
        sh = gc.open_by_url(SPREADSHEET_URL)
        # 取得所有工作表
        all_worksheets = sh.worksheets()
        sheet_names = [ws.title for ws in all_worksheets]
        
        # 選擇工作表
        if sheet_names:
            selected_sheet_name = st.selectbox("📂 選擇工作表 (資料來源)", sheet_names, index=0)
        else:
            st.error("Google Sheet 中沒有任何工作表！")
            st.stop()
            
    except Exception as e:
        st.error(f"連線失敗: {e}")
        st.stop()

    # 讀取/重整按鈕
    need_refresh = st.button("🔄 重新整理資料")
    
    # 自動載入邏輯
    if need_refresh or st.session_state.df_main is None or st.session_state.current_sheet != selected_sheet_name:
        with st.spinner(f"正在讀取「{selected_sheet_name}」..."):
            st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
            st.session_state.current_sheet = selected_sheet_name
            if need_refresh: st.success("資料已更新！")

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名才能開始操作。")
    st.stop()

df = st.session_state.df_main

# 取得目前操作的 worksheet 物件
try:
    worksheet = sh.worksheet(selected_sheet_name)
except:
    # 如果工作表剛剛被刪除了，這裡會報錯，強制重整
    st.warning("工作表可能已被刪除，正在重新整理...")
    st.cache_data.clear()
    time.sleep(1)
    st.rerun()

# ================= 分頁功能 =================
tab_upload, tab_prepare, tab_confirm, tab_manage = st.tabs([
    "📥 建立新工作表", 
    "📄 [1] 準備匯出", 
    "✅ [2] 確認領取", 
    "🛠️ 資料管理 (刪除)"
])

# -------------------------------------------
# TAB 1: 上載新資料 (建立新 Sheet)
# -------------------------------------------
with tab_upload:
    st.subheader("📥 上傳 Excel 並建立獨立工作表")
    
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=['xlsx', 'xls'], key="uploader")
    
    # 使用 session_state 來綁定輸入框，方便清空
    new_sheet_name = st.text_input("請輸入新工作表名稱", placeholder="例如: 2024_第一期", key="sheet_input_key")
    
    if st.button("🚀 建立新表並上傳", type="primary"):
        if not uploaded_file:
            st.error("請先選擇檔案！")
        elif not new_sheet_name:
            st.error("請輸入新工作表名稱！")
        elif new_sheet_name in sheet_names:
            st.error(f"⚠️ 名稱「{new_sheet_name}」已存在，請使用不同名稱。")
        else:
            try:
                new_df = pd.read_excel(uploaded_file)
                
                # 欄位檢查與對應
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
                    
                    # 補上系統欄位 (這就是之前報錯的地方，現在修復了)
                    for col in SYSTEM_COLS:
                        new_df[col] = ''
                    
                    new_df['ID序號'] = new_df['ID序號'].astype(str)
                    new_df = new_df.fillna('')
                    
                    with st.spinner("正在建立新工作表..."):
                        new_ws = sh.add_worksheet(title=new_sheet_name, rows=len(new_df)+50, cols=20)
                        data_to_write = [new_df.columns.tolist()] + new_df.values.tolist()
                        new_ws.update('A1', data_to_write)
                        
                        st.success(f"成功建立「{new_sheet_name}」！")
                        
                        # 清空輸入狀態，避免重複觸發
                        # 注意：Streamlit 不允許直接修改 widget key 的 session state，我們透過重整解決
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()
                else:
                    st.error("上傳的 Excel 欄位不足 9 欄。")
            except Exception as e:
                st.error(f"發生錯誤: {e}")

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
                st.warning("未選取人員")
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
                    
                    time.sleep(2)
                    st.cache_data.clear()
                    st.rerun()

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
                        except: pass
                        prog.progress((i + 1) / len(selected))
                    
                    st.success("更新完成！")
                    time.sleep(1)
                    st.cache_data.clear()
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
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()

# -------------------------------------------
# TAB 4: 資料管理 (刪除功能 - 增強版)
# -------------------------------------------
with tab_manage:
    st.subheader(f"🛠️ 資料管理 - {selected_sheet_name}")
    st.error("⚠️ 警告：此處的操作將直接修改 Google Sheets，且無法復原！")
    
    # 顯示所有資料供勾選
    df_manage = df.copy()
    df_manage.insert(0, "刪除", False)
    
    edited_manage = st.data_editor(
        df_manage,
        column_config={"刪除": st.column_config.CheckboxColumn(required=True, label="選取刪除")},
        hide_index=True,
        key="editor_manage"
    )
    
    st.divider()
    col_d1, col_d2, col_d3 = st.columns(3)
    
    # 功能 1: 刪除選取
    with col_d1:
        st.markdown("##### 🗑️ 刪除選取的列")
        if st.button("執行刪除 (Selected Rows)"):
            selected_del = edited_manage[edited_manage["刪除"] == True]
            if selected_del.empty:
                st.warning("請先勾選上方的資料。")
            else:
                if st.checkbox(f"確定刪除 {len(selected_del)} 筆資料？", key="chk_del_rows"):
                    with st.spinner("刪除中..."):
                        # 收集要刪除的 Row Index
                        rows_to_delete = []
                        for idx, row in selected_del.iterrows():
                            try:
                                cell = worksheet.find(row['ID序號'], in_column=1)
                                if cell:
                                    rows_to_delete.append(cell.row)
                            except: pass
                        
                        # 由大到小排序，避免刪除後 Index 跑掉
                        rows_to_delete.sort(reverse=True)
                        
                        for r_idx in rows_to_delete:
                            worksheet.delete_rows(r_idx)
                            
                        st.success("刪除完成！")
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()

    # 功能 2: 清空整表內容
    with col_d2:
        st.markdown("##### 🧹 清空所有內容 (保留標題)")
        if st.button("執行清空 (Clear Content)"):
            if st.checkbox("確定清空整張表的內容？", key="chk_clear_all"):
                with st.spinner("清空中..."):
                    headers = worksheet.row_values(1)
                    worksheet.clear()
                    # 重新寫入標題
                    worksheet.append_row(headers)
                    st.success("已清空，保留標題列。")
                    time.sleep(1)
                    st.cache_data.clear()
                    st.rerun()

    # 功能 3: 刪除整張工作表 (Sheet)
    with col_d3:
        st.markdown("##### 🔥 刪除整個工作表 (Delete Sheet)")
        if st.button("執行刪除 (Delete Worksheet)", type="primary"):
            # 檢查是否只剩一張表 (Google Sheet 不允許刪除最後一張表)
            if len(sheet_names) <= 1:
                st.error("無法刪除：Google Sheets 至少必須保留一個工作表。")
            else:
                if st.checkbox(f"確定要永久刪除「{selected_sheet_name}」？", key="chk_del_sheet"):
                    with st.spinner(f"正在刪除 {selected_sheet_name}..."):
                        sh.del_worksheet(worksheet)
                        st.success(f"工作表「{selected_sheet_name}」已刪除。")
                        time.sleep(2)
                        st.cache_data.clear()
                        st.rerun()
