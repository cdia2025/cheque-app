import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import io
import time
import re
import os

# ================= 設定區 =================
# Google Sheet ID (從網址提取，比 URL 連線更穩定)
SPREADSHEET_ID = "1gpq9Cye25rmPgyOt508L1sBvlIpPis45R09vn0uy434"
SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit"

# Word 範本檔案名稱
TEMPLATE_FILENAME = "表格二津貼簽收記錄.docx"
TEMPLATE_FILENAME_ENG = "表格二津貼簽收記錄(Eng).docx"

# 系統欄位與順序
REQUIRED_COLS = [
    'ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', 
    '反思會', '反思表', '家長/監護人', 
    'Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff'
]

st.set_page_config(page_title="雲端實習津貼系統 (V64 完整修復版)", layout="wide", page_icon="🛡️")

# ================= 連線設定 =================

# 1. 資料讀寫連線 (Streamlit 官方套件)
conn = st.connection("gsheets", type=GSheetsConnection)

# 2. 結構管理連線 (原生 gspread)
@st.cache_resource
def get_manager_client():
    """建立 gspread 客戶端"""
    try:
        creds_dict = dict(st.secrets["connections"]["gsheets"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"管理連線失敗: {e}")
        st.stop()

# ================= 核心函式 =================

def clean_dataframe(df):
    """資料清洗與格式統一"""
    for col in REQUIRED_COLS:
        if col not in df.columns:
            df[col] = ""
    df = df[REQUIRED_COLS]
    df = df.astype(str)
    for col in df.columns:
        df[col] = df[col].replace(['NaT', 'nan', 'None', '<NA>'], '')
        df[col] = df[col].str.strip()
    df['ID序號'] = df['ID序號'].apply(lambda x: x[:-2] if x.endswith('.0') else x)
    return df

def get_all_sheet_names():
    """取得所有工作表名稱 (加入自動重試機制以解決 500 錯誤)"""
    client = get_manager_client()
    max_retries = 3
    
    for attempt in range(max_retries):
        try:
            # 改用 open_by_key 比較穩定
            sh = client.open_by_key(SPREADSHEET_ID)
            return [ws.title for ws in sh.worksheets()]
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(2) # 等待 2 秒後重試
                continue
            else:
                st.error(f"無法讀取工作表清單 (Google 伺服器忙碌): {e}")
                return []

def load_data(sheet_name):
    """讀取資料"""
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl=0)
        return clean_dataframe(df)
    except:
        return pd.DataFrame(columns=REQUIRED_COLS)

def save_data(df, sheet_name):
    """儲存資料"""
    try:
        clean_df = clean_dataframe(df)
        conn.update(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, data=clean_df)
        st.toast("✅ 資料已同步！", icon="☁️")
        st.session_state.df_main = clean_df
        return True
    except Exception as e:
        if "429" in str(e):
            st.error("⚠️ 流量過大，請稍後再試。")
        else:
            st.error(f"儲存失敗: {e}")
        return False

def delete_worksheet(worksheet_name):
    """刪除指定的工作表"""
    try:
        client = get_manager_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        ws = sh.worksheet(worksheet_name)
        sh.del_worksheet(ws)
        
        if st.session_state.current_sheet == worksheet_name:
            sheet_names = get_all_sheet_names()
            if sheet_names:
                st.session_state.current_sheet = sheet_names[0]
                st.session_state.df_main = load_data(st.session_state.current_sheet)
            else:
                st.session_state.current_sheet = None
                st.session_state.df_main = None
        
        st.success(f"工作表 '{worksheet_name}' 已刪除")
        return True
    except Exception as e:
        st.error(f"刪除工作表失敗: {e}")
        return False

def calculate_statistics(df):
    """計算統計數字"""
    total_count = len(df)
    ready_for_export = len(df[(df['反思會'].str.upper() == 'Y') & (df['反思表'].str.upper() == 'Y') & (df['DocGeneratedDate'] == '')])
    pending_collection = len(df[(df['DocGeneratedDate'] != '') & (df['Collected'] != 'Y')])
    collected = len(df[df['Collected'] == 'Y'])
    not_qualified = len(df[((df['反思會'].str.upper() != 'Y') | (df['反思表'].str.upper() != 'Y')) & (df['DocGeneratedDate'] == '')])
    
    return {
        'total': total_count,
        'ready_for_export': ready_for_export,
        'pending_collection': pending_collection,
        'collected': collected,
        'not_qualified': not_qualified
    }

def process_batch_selection(df_target, check_col_name, key_suffix):
    """批量選取邏輯"""
    ss_select_all = f"select_all_{key_suffix}"
    
    if ss_select_all not in st.session_state:
        st.session_state[ss_select_all] = False

    if check_col_name not in df_target.columns:
        df_target.insert(0, check_col_name, False)

    with st.expander("⚡ 批量選取工具 (輸入 ID 或 全選)", expanded=False):
        c1, c2 = st.columns([3, 1])
        with c1:
            batch_text = st.text_area(
                "貼上 ID (支援 Excel 複製貼上、逗號或空白分隔)", 
                height=100, 
                key=f"batch_txt_{key_suffix}",
                placeholder="例如：\n112001\n112005\n112008"
            )
        with c2:
            st.write("快捷鍵")
            if st.button("✅ 全選列表", key=f"all_{key_suffix}"):
                st.session_state[ss_select_all] = True
            
            if st.button("❌ 全部取消", key=f"clear_{key_suffix}"):
                st.session_state[ss_select_all] = False

        if st.session_state[ss_select_all]:
            df_target[check_col_name] = True
            st.caption("🔴 目前狀態：全選模式")
            
        elif batch_text:
            ids_input = re.split(r'[,\s\n\t]+', batch_text)
            ids_input = [x.strip() for x in ids_input if x.strip()]
            
            if ids_input:
                mask = df_target['ID序號'].isin(ids_input)
                df_target.loc[mask, check_col_name] = True
                match_count = mask.sum()
                st.caption(f"已選取 {match_count} 筆符合的資料")

    return df_target

def perform_global_search(query):
    """執行全域搜尋 (搜尋所有工作表)"""
    results = []
    all_sheets = get_all_sheet_names()
    
    progress_bar = st.progress(0, text="準備開始搜尋...")
    total_sheets = len(all_sheets)
    
    for i, sheet_name in enumerate(all_sheets):
        progress_bar.progress((i + 1) / total_sheets, text=f"正在搜尋工作表：{sheet_name} ({i+1}/{total_sheets})")
        
        try:
            df_temp = load_data(sheet_name)
            if df_temp.empty: continue

            search_cols = ['ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話']
            valid_cols = [c for c in search_cols if c in df_temp.columns]
            
            mask = df_temp[valid_cols].astype(str).apply(
                lambda x: x.str.contains(query, case=False, na=False)
            ).any(axis=1)
            
            found_rows = df_temp[mask]
            
            for _, row in found_rows.iterrows():
                status = "未知"
                if row['Collected'] == 'Y': status = "🟢 已取票"
                elif row['DocGeneratedDate'] != '': status = "🔵 待領取"
                elif row['反思會'].upper() == 'Y' and row['反思表'].upper() == 'Y': status = "📄 準備匯出"
                else: status = "🚫 不符/其他"

                results.append({
                    "來源工作表": sheet_name,
                    "ID序號": row['ID序號'],
                    "姓名(中文)": row['姓名(中文)'],
                    "電話": row['電話'],
                    "目前狀態": status,
                    "DocDate": row['DocGeneratedDate']
                })
        except Exception as e:
            print(f"搜尋 {sheet_name} 時發生錯誤: {e}")
            
    progress_bar.empty()
    return pd.DataFrame(results)

# ================= Session State =================
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None
if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'export_file' not in st.session_state: st.session_state.export_file = None
if 'staff_name' not in st.session_state: st.session_state.staff_name = ""
if 'show_delete_confirmation' not in st.session_state: 
    st.session_state.show_delete_confirmation = False
    st.session_state.delete_sheet_name = ""
if 'search_results' not in st.session_state: st.session_state.search_results = None

# ================= 側邊欄 =================
with st.sidebar:
    st.header("LayoutPanel")
    staff_name = st.text_input("👤 負責職員姓名", value=st.session_state.get('staff_name', ''), key="staff_name_input")
    if staff_name: st.session_state.staff_name = staff_name
    
    st.divider()
    
    sheet_names = get_all_sheet_names()
    if not sheet_names: st.stop()
        
    if st.session_state.current_sheet not in sheet_names:
        st.session_state.current_sheet = sheet_names[0]
        
    idx = sheet_names.index(st.session_state.current_sheet)
    selected_sheet = st.selectbox("📂 選擇工作表", sheet_names, index=idx)
    
    if selected_sheet != st.session_state.current_sheet:
        st.session_state.current_sheet = selected_sheet
        st.session_state.df_main = load_data(selected_sheet)
        st.session_state.export_file = None
        for key in list(st.session_state.keys()):
            if key.startswith("select_all_"): st.session_state[key] = False
        st.rerun()

    st.divider()

    st.subheader("📂 下載合併範本")
    if os.path.exists(TEMPLATE_FILENAME):
        with open(TEMPLATE_FILENAME, "rb") as f:
            st.download_button("📥 下載：表格二津貼簽收記錄", f, TEMPLATE_FILENAME, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_btn_zh")
    else:
        st.info(f"💡 缺少檔案: {TEMPLATE_FILENAME}")

    if os.path.exists(TEMPLATE_FILENAME_ENG):
        with open(TEMPLATE_FILENAME_ENG, "rb") as f:
            st.download_button("📥 下載：表格二津貼簽收記錄 (Eng)", f, TEMPLATE_FILENAME_ENG, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_btn_eng")
    else:
        st.info(f"💡 缺少檔案: {TEMPLATE_FILENAME_ENG}")

    st.divider()

    st.subheader("🗑️ 管理工作表")
    delete_sheet = st.selectbox("選擇要刪除的工作表", [""] + [name for name in sheet_names if name != selected_sheet])
    
    if delete_sheet:
        if st.button(f"🗑️ 刪除工作表 '{delete_sheet}'", type="secondary"):
            st.session_state.show_delete_confirmation = True
            st.session_state.delete_sheet_name = delete_sheet
    
    if st.session_state.show_delete_confirmation:
        st.warning(f"⚠️ 確定要永久刪除工作表 '{st.session_state.delete_sheet_name}' 嗎？")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ 確定刪除", type="primary"):
                if delete_worksheet(st.session_state.delete_sheet_name):
                    st.session_state.show_delete_confirmation = False
                    st.session_state.delete_sheet_name = ""
                    st.rerun()
        with c2:
            if st.button("❌ 取消"):
                st.session_state.show_delete_confirmation = False
                st.session_state.delete_sheet_name = ""
                st.rerun()

    if st.button("🔄 強制重新整理"):
        st.cache_data.clear()
        st.session_state.df_main = load_data(selected_sheet)
        st.session_state.export_file = None
        st.session_state.search_results = None
        for key in list(st.session_state.keys()):
            if key.startswith("select_all_"): st.session_state[key] = False
        st.rerun()

if not staff_name:
    st.warning("⚠️ 請先在左側輸入姓名。")
    st.stop()

if st.session_state.df_main is None:
    st.session_state.df_main = load_data(selected_sheet)

df = st.session_state.df_main
st.title(f"☁️ 管理：{selected_sheet}")

stats = calculate_statistics(df)
col1, col2, col3, col4, col5 = st.columns(5)
with col1: st.metric("📊 總人數", stats['total'])
with col2: st.metric("📄 準備匯出", stats['ready_for_export'])
with col3: st.metric("🔵 待領取", stats['pending_collection'])
with col4: st.metric("🟢 已取票", stats['collected'])
with col5: st.metric("🚫 不符", stats['not_qualified'])

st.divider()

# ================= 主分頁 (改用 Radio Button 當作導覽列以防止跳頁) =================
PAGES = [
    "📥 建立/上傳", 
    "📄 [1] 準備匯出", 
    "🔵 [2] 待領取", 
    "🟢 [3] 已取票", 
    "🚫 [4] 不符", 
    "✏️ 修改", 
    "🔍 全域搜尋"
]

selected_page = st.radio(
    "導覽選單", 
    PAGES, 
    horizontal=True, 
    label_visibility="collapsed",
    key="nav_radio"
)

st.divider()

# ---------------- 頁面邏輯 ----------------

if selected_page == "📥 建立/上傳":
    st.subheader("上傳 Excel 並建立新分頁")
    up_file = st.file_uploader("選擇 Excel", type=["xlsx", "xls"], key="upload_tab1")
    new_name = st.text_input("新工作表名稱 (如: 2024_05)", key="new_name_tab1")
    if st.button("🚀 建立並上傳", type="primary"):
        if up_file and new_name:
            if new_name in sheet_names:
                st.error("名稱重複！")
            else:
                try:
                    new_df = pd.read_excel(up_file)
                    if len(new_df.columns) >= 9:
                        mapping = {
                            new_df.columns[0]: 'ID序號', new_df.columns[1]: '編號',
                            new_df.columns[2]: '姓名(中文)', new_df.columns[3]: '姓名(英文)',
                            new_df.columns[4]: '電話', new_df.columns[5]: '實習日數',
                            new_df.columns[6]: '反思會', new_df.columns[7]: '反思表',
                            new_df.columns[8]: '家長/監護人'
                        }
                        new_df.rename(columns=mapping, inplace=True)
                        for c in ['Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']:
                            new_df[c] = ""
                        
                        client = get_manager_client()
                        sh = client.open_by_key(SPREADSHEET_ID)
                        ws = sh.add_worksheet(title=new_name, rows=len(new_df)+20, cols=15)
                        
                        clean_new = clean_dataframe(new_df)
                        data_export = [clean_new.columns.tolist()] + clean_new.values.tolist()
                        ws.update(data_export)
                        
                        st.success("建立成功！")
                        st.session_state.current_sheet = new_name
                        time.sleep(2)
                        st.rerun()
                    else:
                        st.error("欄位不足")
                except Exception as e:
                    st.error(f"錯誤: {e}")

elif selected_page == "📄 [1] 準備匯出":
    st.subheader("步驟一：匯出資料")
    if st.session_state.export_file:
        st.success("✅ 匯出成功！")
        st.download_button("📥 下載 MailMerge Source", st.session_state.export_file, "MailMerge_Source.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        st.divider()

    mask = (df['反思會'].str.upper() == 'Y') & (df['反思表'].str.upper() == 'Y') & (df['DocGeneratedDate'] == '')
    df_show = df[mask].copy()
    df_show = process_batch_selection(df_show, "選取", "tab2")
    
    edited = st.data_editor(df_show, column_config={"選取": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df_show.columns if c != "選取"], hide_index=True)
    
    if st.button("📤 匯出 & 更新狀態"):
        selected = edited[edited["選取"]]
        if selected.empty:
            st.warning("未選取")
        else:
            today = datetime.now().strftime("%Y-%m-%d")
            ids = selected['ID序號'].tolist()
            df.loc[df['ID序號'].isin(ids), 'DocGeneratedDate'] = today
            df.loc[df['ID序號'].isin(ids), 'ResponsibleStaff'] = staff_name
            if save_data(df, selected_sheet):
                st.session_state["select_all_tab2"] = False
                out_df = selected.drop(columns=['選取'])
                out_df['StaffName'] = staff_name
                out_df['TodayDate'] = today
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: out_df.to_excel(writer, index=False)
                st.session_state.export_file = buffer.getvalue()
                st.rerun()

elif selected_page == "🔵 [2] 待領取":
    st.subheader("步驟二：準備領取")
    mask = (df['DocGeneratedDate'] != '') & (df['Collected'] != 'Y')
    df_show = df[mask].copy()
    df_show = process_batch_selection(df_show, "確認", "tab3")
    edited = st.data_editor(df_show, column_config={"確認": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df_show.columns if c != "確認"], hide_index=True)
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ 確認已取票", type="primary"):
            ids = edited[edited["確認"]]['ID序號'].tolist()
            if ids:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                df.loc[df['ID序號'].isin(ids), 'Collected'] = 'Y'
                df.loc[df['ID序號'].isin(ids), 'CollectedDate'] = now
                save_data(df, selected_sheet)
                st.session_state["select_all_tab3"] = False
                st.rerun()
    with c2:
        if st.button("↩️ 退回"):
            ids = edited[edited["確認"]]['ID序號'].tolist()
            if ids:
                df.loc[df['ID序號'].isin(ids), 'DocGeneratedDate'] = ''
                df.loc[df['ID序號'].isin(ids), 'ResponsibleStaff'] = ''
                save_data(df, selected_sheet)
                st.session_state["select_all_tab3"] = False
                st.rerun()

elif selected_page == "🟢 [3] 已取票":
    st.subheader("已取票紀錄")
    mask = (df['Collected'] == 'Y')
    df_show = df[mask].copy()
    df_show = process_batch_selection(df_show, "撤銷", "tab4")
    edited = st.data_editor(df_show, column_config={"撤銷": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df_show.columns if c != "撤銷"], hide_index=True)
    if st.button("↩️ 撤銷領取"):
        ids = edited[edited["撤銷"]]['ID序號'].tolist()
        if ids:
            df.loc[df['ID序號'].isin(ids), 'Collected'] = ''
            df.loc[df['ID序號'].isin(ids), 'CollectedDate'] = ''
            save_data(df, selected_sheet)
            st.session_state["select_all_tab4"] = False
            st.rerun()

elif selected_page == "🚫 [4] 不符":
    st.subheader("不符資格名單")
    mask = ((df['反思會'].str.upper() != 'Y') | (df['反思表'].str.upper() != 'Y')) & (df['DocGeneratedDate'] == '')
    df_show = df[mask].copy()
    df_show = process_batch_selection(df_show, "放行", "tab5")
    edited = st.data_editor(df_show, column_config={"放行": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df_show.columns if c != "放行"], hide_index=True)
    if st.button("➡️ 強制放行"):
        ids = edited[edited["放行"]]['ID序號'].tolist()
        if ids:
            df.loc[df['ID序號'].isin(ids), '反思會'] = 'Y'
            df.loc[df['ID序號'].isin(ids), '反思表'] = 'Y'
            save_data(df, selected_sheet)
            st.session_state["select_all_tab5"] = False
            st.rerun()

elif selected_page == "✏️ 修改":
    st.subheader("✏️ 直接編輯")
    df_edit = df.copy()
    edited_df = st.data_editor(
        df_edit,
        column_config={
            "反思會": st.column_config.SelectboxColumn(options=["Y", "N", ""], required=True),
            "反思表": st.column_config.SelectboxColumn(options=["Y", "N", ""], required=True),
            "實習日數": st.column_config.NumberColumn(min_value=0, max_value=365, step=1),
        },
        disabled=['ID序號', 'Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff'],
        hide_index=True,
        width='stretch',
        key="editor_main"
    )
    if st.button("💾 儲存全部修改", type="primary"):
        save_data(edited_df, selected_sheet)
        st.rerun()

elif selected_page == "🔍 全域搜尋":
    st.subheader("🔍 搜尋全系統資料")
    col_search, col_btn = st.columns([4, 1])
    with col_search:
        search_query = st.text_input("輸入關鍵字 (ID、姓名或電話)")
    with col_btn:
        st.write("")
        st.write("")
        if st.button("🚀 開始搜尋", type="primary"):
            if not search_query: st.warning("請輸入關鍵字")
            else: st.session_state.search_results = perform_global_search(search_query)

    st.divider()

    if st.session_state.search_results is not None:
        if st.session_state.search_results.empty:
            st.warning("❌ 未找到資料")
        else:
            st.success(f"✅ 找到 {len(st.session_state.search_results)} 筆：")
            st.dataframe(
                st.session_state.search_results,
                column_config={"來源工作表": st.column_config.TextColumn("位於工作表"), "DocDate": st.column_config.TextColumn("匯出日期")},
                width='stretch',
                hide_index=True
            )
