import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import io
import time

# ================= 設定區 =================
# 請確認這裡是您的 Google Sheet 網址
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1gpq9Cye25rmPgyOt508L1sBvlIpPis45R09vn0uy434/edit"

# 系統欄位與順序
REQUIRED_COLS = [
    'ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', 
    '反思會', '反思表', '家長/監護人', 
    'Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff'
]

st.set_page_config(page_title="雲端實習津貼系統 (V53 連線修復版)", layout="wide", page_icon="🛡️")

# ================= 連線設定 =================

# 1. 資料讀寫連線 (使用 Streamlit 官方套件)
conn = st.connection("gsheets", type=GSheetsConnection)

# 2. 結構管理連線 (使用原生 gspread，修復 open_by_url 錯誤)
@st.cache_resource
def get_manager_client():
    """建立一個原生的 gspread 客戶端，用於管理工作表結構"""
    try:
        # 從 secrets 讀取
        creds_dict = dict(st.secrets["connections"]["gsheets"])
        
        # 修正私鑰換行問題
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            
        # 定義權限範圍
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        
        # 建立憑證
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"管理連線失敗: {e}")
        st.stop()

# ================= 核心函式 =================

def clean_dataframe(df):
    """資料清洗與格式統一"""
    # 補齊欄位
    for col in REQUIRED_COLS:
        if col not in df.columns:
            df[col] = ""
    # 排序欄位
    df = df[REQUIRED_COLS]
    # 轉字串
    df = df.astype(str)
    # 清理內容
    for col in df.columns:
        df[col] = df[col].replace(['NaT', 'nan', 'None', '<NA>'], '')
        df[col] = df[col].str.strip()
    # 處理 ID
    df['ID序號'] = df['ID序號'].apply(lambda x: x[:-2] if x.endswith('.0') else x)
    return df

def get_all_sheet_names():
    """取得所有工作表名稱 (使用 manager client)"""
    try:
        client = get_manager_client()
        sh = client.open_by_url(SPREADSHEET_URL)
        return [ws.title for ws in sh.worksheets()]
    except Exception as e:
        st.error(f"無法讀取工作表清單: {e}")
        return []

def load_data(sheet_name):
    """讀取資料 (使用 conn)"""
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl=0)
        return clean_dataframe(df)
    except:
        return pd.DataFrame(columns=REQUIRED_COLS)

def save_data(df, sheet_name):
    """儲存資料 (使用 conn update)"""
    try:
        clean_df = clean_dataframe(df)
        conn.update(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, data=clean_df)
        st.toast("✅ 資料已同步！", icon="☁️")
        st.session_state.df_main = clean_df # 更新本地快取
        return True
    except Exception as e:
        if "429" in str(e):
            st.error("⚠️ 流量過大，請稍後再試。")
        else:
            st.error(f"儲存失敗: {e}")
        return False

def select_all_rows(df, selection_column, select=True):
    """全選或取消全選指定欄位"""
    df_copy = df.copy()
    df_copy[selection_column] = select
    return df_copy

# ================= Session State =================
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None
if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'export_file' not in st.session_state: st.session_state.export_file = None
if 'staff_name' not in st.session_state: st.session_state.staff_name = ""

# ================= 側邊欄 =================
with st.sidebar:
    st.header("LayoutPanel")
    staff_name = st.text_input("👤 負責職員姓名", value=st.session_state.get('staff_name', ''), key="staff_name_input")
    
    # 更新session state
    if staff_name:
        st.session_state.staff_name = staff_name
    
    st.divider()
    
    # 1. 取得工作表清單
    sheet_names = get_all_sheet_names()
    if not sheet_names:
        st.stop()
        
    # 2. 選擇工作表 (鎖定 Index)
    if st.session_state.current_sheet not in sheet_names:
        st.session_state.current_sheet = sheet_names[0]
        
    idx = sheet_names.index(st.session_state.current_sheet)
    selected_sheet = st.selectbox("📂 選擇工作表", sheet_names, index=idx)
    
    # 切換檢測
    if selected_sheet != st.session_state.current_sheet:
        st.session_state.current_sheet = selected_sheet
        st.session_state.df_main = load_data(selected_sheet)
        st.session_state.export_file = None
        st.rerun()

    if st.button("🔄 強制重新整理"):
        st.cache_data.clear()
        st.session_state.df_main = load_data(selected_sheet)
        st.session_state.export_file = None
        st.rerun()

if not staff_name:
    st.warning("⚠️ 請先在左側輸入姓名。")
    st.stop()

if st.session_state.df_main is None:
    st.session_state.df_main = load_data(selected_sheet)

df = st.session_state.df_main
st.title(f"☁️ 管理：{selected_sheet}")

# ================= 主分頁 =================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📥 建立/上傳", "📄 [1] 準備匯出", "🔵 [2] 待領取", "🟢 [3] 已取票", "🚫 [4] 不符", "✏️ 修改"
])

# ---------------- TAB 1: 建立新表 ----------------
with tab1:
    st.subheader("上傳 Excel 並建立新分頁")
    up_file = st.file_uploader("選擇 Excel", type=["xlsx", "xls"], key="upload_tab1")
    new_name = st.text_input("新工作表名稱 (如: 2024_05)", key="new_name_tab1")
    
    if st.button("🚀 建立並上傳", type="primary", key="create_upload_btn"):
        if up_file and new_name:
            if new_name in sheet_names:
                st.error("名稱重複！")
            else:
                try:
                    new_df = pd.read_excel(up_file)
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
                        for c in ['Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']:
                            new_df[c] = ""
                        
                        # 使用 manager client 建立
                        client = get_manager_client()
                        sh = client.open_by_url(SPREADSHEET_URL)
                        ws = sh.add_worksheet(title=new_name, rows=len(new_df)+20, cols=15)
                        
                        # 寫入資料
                        clean_new = clean_dataframe(new_df)
                        # gspread update 需要 list of lists
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

# ---------------- TAB 2: 準備匯出 ----------------
with tab2:
    st.subheader("步驟一：匯出資料")
    
    if st.session_state.export_file:
        st.success("✅ 匯出成功！請下載：")
        st.download_button("📥 下載 MailMerge Source", st.session_state.export_file, "MailMerge_Source.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        st.divider()

    mask = (df['反思會'].str.upper() == 'Y') & (df['反思表'].str.upper() == 'Y') & (df['DocGeneratedDate'] == '')
    df_show = df[mask].copy()
    
    # 添加批量選取按鈕
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("✅ 全選", key="select_all_tab2"):
            df_show = select_all_rows(df_show, "選取", True)
        if st.button("❌ 取消全選", key="deselect_all_tab2"):
            df_show = select_all_rows(df_show, "選取", False)
    
    df_show.insert(0, "選取", False)
    edited = st.data_editor(
        df_show, 
        column_config={"選取": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df_show.columns if c != "選取"],
        hide_index=True,
        key="editor_tab2"
    )
    
    if st.button("📤 匯出 & 更新狀態", key="export_status_btn"):
        selected = edited[edited["選取"]]
        if selected.empty:
            st.warning("未選取任何項目")
        else:
            today = datetime.now().strftime("%Y-%m-%d")
            ids = selected['ID序號'].tolist()
            
            # Pandas 更新
            df.loc[df['ID序號'].isin(ids), 'DocGeneratedDate'] = today
            df.loc[df['ID序號'].isin(ids), 'ResponsibleStaff'] = staff_name
            
            # 存雲端
            if save_data(df, selected_sheet):
                out_df = selected.drop(columns=['選取'])
                out_df['StaffName'] = staff_name
                out_df['TodayDate'] = today
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    out_df.to_excel(writer, index=False)
                
                st.session_state.export_file = buffer.getvalue()
                st.rerun()

# ---------------- TAB 3: 待領取 ----------------
with tab3:
    st.subheader("步驟二：準備領取")
    mask = (df['DocGeneratedDate'] != '') & (df['Collected'] != 'Y')
    df_show = df[mask].copy()
    
    # 添加批量選取按鈕
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("✅ 全選", key="select_all_tab3"):
            df_show = select_all_rows(df_show, "確認", True)
        if st.button("❌ 取消全選", key="deselect_all_tab3"):
            df_show = select_all_rows(df_show, "確認", False)
    
    df_show.insert(0, "確認", False)
    edited = st.data_editor(
        df_show, 
        column_config={"確認": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df_show.columns if c != "確認"],
        hide_index=True,
        key="editor_tab3"
    )
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ 確認已取票", type="primary", key="confirm_collected_btn"):
            ids = edited[edited["確認"]]['ID序號'].tolist()
            if ids:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                df.loc[df['ID序號'].isin(ids), 'Collected'] = 'Y'
                df.loc[df['ID序號'].isin(ids), 'CollectedDate'] = now
                save_data(df, selected_sheet)
                st.rerun()
    with c2:
        if st.button("↩️ 退回至準備匯出", key="revert_to_export_btn"):
            ids = edited[edited["確認"]]['ID序號'].tolist()
            if ids:
                df.loc[df['ID序號'].isin(ids), 'DocGeneratedDate'] = ''
                df.loc[df['ID序號'].isin(ids), 'ResponsibleStaff'] = ''
                save_data(df, selected_sheet)
                st.rerun()

# ---------------- TAB 4: 已取票 ----------------
with tab4:
    st.subheader("已取票紀錄")
    mask = (df['Collected'] == 'Y')
    df_show = df[mask].copy()
    
    # 添加批量選取按鈕
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("✅ 全選", key="select_all_tab4"):
            df_show = select_all_rows(df_show, "撤銷", True)
        if st.button("❌ 取消全選", key="deselect_all_tab4"):
            df_show = select_all_rows(df_show, "撤銷", False)
    
    df_show.insert(0, "撤銷", False)
    edited = st.data_editor(
        df_show, 
        column_config={"撤銷": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df_show.columns if c != "撤銷"],
        hide_index=True,
        key="editor_tab4"
    )
    
    if st.button("↩️ 撤銷領取", key="revert_collected_btn"):
        ids = edited[edited["撤銷"]]['ID序號'].tolist()
        if ids:
            df.loc[df['ID序號'].isin(ids), 'Collected'] = ''
            df.loc[df['ID序號'].isin(ids), 'CollectedDate'] = ''
            save_data(df, selected_sheet)
            st.rerun()

# ---------------- TAB 5: 不符名單 ----------------
with tab5:
    st.subheader("不符資格名單")
    mask = ((df['反思會'].str.upper() != 'Y') | (df['反思表'].str.upper() != 'Y')) & (df['DocGeneratedDate'] == '')
    df_show = df[mask].copy()
    
    # 添加批量選取按鈕
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("✅ 全選", key="select_all_tab5"):
            df_show = select_all_rows(df_show, "放行", True)
        if st.button("❌ 取消全選", key="deselect_all_tab5"):
            df_show = select_all_rows(df_show, "放行", False)
    
    df_show.insert(0, "放行", False)
    edited = st.data_editor(
        df_show, 
        column_config={"放行": st.column_config.CheckboxColumn(required=True)},
        disabled=[c for c in df_show.columns if c != "放行"],
        hide_index=True,
        key="editor_tab5"
    )
    
    if st.button("➡️ 強制放行", key="force_approve_btn"):
        ids = edited[edited["放行"]]['ID序號'].tolist()
        if ids:
            df.loc[df['ID序號'].isin(ids), '反思會'] = 'Y'
            df.loc[df['ID序號'].isin(ids), '反思表'] = 'Y'
            save_data(df, selected_sheet)
            st.rerun()

# ---------------- TAB 6: 修改資料 ----------------
with tab6:
    st.subheader("✏️ 直接編輯")
    st.info("直接修改，完成後按「儲存」。")
    
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
        use_container_width=True,
        key="editor_main"
    )
    
    if st.button("💾 儲存全部修改", type="primary", key="save_all_changes_btn"):
        save_data(edited_df, selected_sheet)
        st.rerun()
