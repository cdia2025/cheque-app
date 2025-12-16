import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import io
import time
import re

# ================= 設定區 =================
# 請確認這裡是您的 Google Sheet 網址
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1gpq9Cye25rmPgyOt508L1sBvlIpPis45R09vn0uy434/edit"

# 系統欄位與順序
REQUIRED_COLS = [
    'ID序號', '編號', '姓名(中文)', '姓名(英文)', '電話', '實習日數', 
    '反思會', '反思表', '家長/監護人', 
    'Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff'
]

st.set_page_config(page_title="雲端實習津貼系統 (V55 全選修復版)", layout="wide", page_icon="🛡️")

# ================= 連線設定 =================

# 1. 資料讀寫連線 (使用 Streamlit 官方套件)
conn = st.connection("gsheets", type=GSheetsConnection)

# 2. 結構管理連線 (使用原生 gspread)
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

def delete_worksheet(worksheet_name):
    """刪除指定的工作表"""
    try:
        client = get_manager_client()
        sh = client.open_by_url(SPREADSHEET_URL)
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
    """計算各類別的統計數字"""
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
    """
    處理批量選取邏輯 (修復版：使用 session_state 記憶狀態)
    """
    # 定義 Session State 的 Key
    ss_select_all = f"select_all_{key_suffix}"
    
    # 初始化狀態
    if ss_select_all not in st.session_state:
        st.session_state[ss_select_all] = False

    # 確保欄位存在
    if check_col_name not in df_target.columns:
        df_target.insert(0, check_col_name, False)

    with st.expander("⚡ 批量選取工具 (輸入 ID 或 全選)", expanded=False):
        c1, c2 = st.columns([3, 1])
        with c1:
            # 文字區域本身就會保留狀態，不需要額外處理
            batch_text = st.text_area(
                "貼上 ID (支援 Excel 複製貼上、逗號或空白分隔)", 
                height=100, 
                key=f"batch_txt_{key_suffix}",
                placeholder="例如：\n112001\n112005\n112008"
            )
        with c2:
            st.write("快捷鍵")
            # 按鈕邏輯：點擊後更新 session state
            if st.button("✅ 全選列表", key=f"all_{key_suffix}"):
                st.session_state[ss_select_all] = True
            
            if st.button("❌ 全部取消", key=f"clear_{key_suffix}"):
                st.session_state[ss_select_all] = False
                # 如果需要也可以清除文字框，但 Streamlit 清除文字框比較複雜，先清空全選狀態即可

        # === 核心邏輯：根據狀態修改 DataFrame ===
        
        # 1. 優先處理「全選」狀態
        if st.session_state[ss_select_all]:
            df_target[check_col_name] = True
            st.caption("🔴 目前狀態：全選模式 (若要取消請按「全部取消」)")
            
        # 2. 如果沒有全選，則檢查文字框 ID
        elif batch_text:
            ids_input = re.split(r'[,\s\n\t]+', batch_text)
            ids_input = [x.strip() for x in ids_input if x.strip()]
            
            if ids_input:
                mask = df_target['ID序號'].isin(ids_input)
                df_target.loc[mask, check_col_name] = True
                match_count = mask.sum()
                st.caption(f"已選取 {match_count} 筆符合的資料")
        
        # 3. 預設狀態 (全選為 False 且無文字) -> 保持原樣 (全部 False)

    return df_target

# ================= Session State =================
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None
if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'export_file' not in st.session_state: st.session_state.export_file = None
if 'staff_name' not in st.session_state: st.session_state.staff_name = ""
if 'show_delete_confirmation' not in st.session_state: 
    st.session_state.show_delete_confirmation = False
    st.session_state.delete_sheet_name = ""

# ================= 側邊欄 =================
with st.sidebar:
    st.header("LayoutPanel")
    staff_name = st.text_input("👤 負責職員姓名", value=st.session_state.get('staff_name', ''), key="staff_name_input")
    
    if staff_name:
        st.session_state.staff_name = staff_name
    
    st.divider()
    
    sheet_names = get_all_sheet_names()
    if not sheet_names:
        st.stop()
        
    if st.session_state.current_sheet not in sheet_names:
        st.session_state.current_sheet = sheet_names[0]
        
    idx = sheet_names.index(st.session_state.current_sheet)
    selected_sheet = st.selectbox("📂 選擇工作表", sheet_names, index=idx)
    
    if selected_sheet != st.session_state.current_sheet:
        st.session_state.current_sheet = selected_sheet
        st.session_state.df_main = load_data(selected_sheet)
        st.session_state.export_file = None
        # 切換工作表時重置全選狀態
        for key in list(st.session_state.keys()):
            if key.startswith("select_all_"):
                st.session_state[key] = False
        st.rerun()

    st.divider()
    st.subheader("🗑️ 管理工作表")
    delete_sheet = st.selectbox("選擇要刪除的工作表", [""] + [name for name in sheet_names if name != selected_sheet])
    
    if delete_sheet:
        if st.button(f"🗑️ 刪除工作表 '{delete_sheet}'", type="secondary"):
            st.session_state.show_delete_confirmation = True
            st.session_state.delete_sheet_name = delete_sheet
    
    if st.session_state.show_delete_confirmation:
        st.warning(f"⚠️ 確定要永久刪除工作表 '{st.session_state.delete_sheet_name}' 嗎？此操作無法還原！")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ 確定刪除", type="primary"):
                if delete_worksheet(st.session_state.delete_sheet_name):
                    st.session_state.show_delete_confirmation = False
                    st.session_state.delete_sheet_name = ""
                    st.rerun()
        with col2:
            if st.button("❌ 取消"):
                st.session_state.show_delete_confirmation = False
                st.session_state.delete_sheet_name = ""
                st.rerun()

    if st.button("🔄 強制重新整理"):
        st.cache_data.clear()
        st.session_state.df_main = load_data(selected_sheet)
        st.session_state.export_file = None
        # 重置全選狀態
        for key in list(st.session_state.keys()):
            if key.startswith("select_all_"):
                st.session_state[key] = False
        st.rerun()

if not staff_name:
    st.warning("⚠️ 請先在左側輸入姓名。")
    st.stop()

if st.session_state.df_main is None:
    st.session_state.df_main = load_data(selected_sheet)

df = st.session_state.df_main
st.title(f"☁️ 管理：{selected_sheet}")

# ================= 統計資料顯示 =================
stats = calculate_statistics(df)
col1, col2, col3, col4, col5 = st.columns(5)
with col1: st.metric(label="📊 總人數", value=stats['total'])
with col2: st.metric(label="📄 準備匯出", value=stats['ready_for_export'], delta_color="off")
with col3: st.metric(label="🔵 待領取", value=stats['pending_collection'], delta_color="off")
with col4: st.metric(label="🟢 已取票", value=stats['collected'], delta_color="off")
with col5: st.metric(label="🚫 不符", value=stats['not_qualified'], delta_color="off")

st.divider()

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
                        sh = client.open_by_url(SPREADSHEET_URL)
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

# ---------------- TAB 2: 準備匯出 ----------------
with tab2:
    st.subheader("步驟一：匯出資料")
    
    if st.session_state.export_file:
        st.success("✅ 匯出成功！請下載：")
        st.download_button("📥 下載 MailMerge Source", st.session_state.export_file, "MailMerge_Source.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        st.divider()

    mask = (df['反思會'].str.upper() == 'Y') & (df['反思表'].str.upper() == 'Y') & (df['DocGeneratedDate'] == '')
    df_show = df[mask].copy()
    
    # === 套用批量選取邏輯 ===
    df_show = process_batch_selection(df_show, "選取", "tab2")
    
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
            st.warning("未選取任何資料")
        else:
            today = datetime.now().strftime("%Y-%m-%d")
            ids = selected['ID序號'].tolist()
            
            df.loc[df['ID序號'].isin(ids), 'DocGeneratedDate'] = today
            df.loc[df['ID序號'].isin(ids), 'ResponsibleStaff'] = staff_name
            
            if save_data(df, selected_sheet):
                # 清除全選狀態
                st.session_state["select_all_tab2"] = False
                
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
    
    # === 套用批量選取邏輯 ===
    df_show = process_batch_selection(df_show, "確認", "tab3")
    
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
                st.session_state["select_all_tab3"] = False # 重置全選
                st.rerun()
    with c2:
        if st.button("↩️ 退回至準備匯出", key="revert_to_export_btn"):
            ids = edited[edited["確認"]]['ID序號'].tolist()
            if ids:
                df.loc[df['ID序號'].isin(ids), 'DocGeneratedDate'] = ''
                df.loc[df['ID序號'].isin(ids), 'ResponsibleStaff'] = ''
                save_data(df, selected_sheet)
                st.session_state["select_all_tab3"] = False # 重置全選
                st.rerun()

# ---------------- TAB 4: 已取票 ----------------
with tab4:
    st.subheader("已取票紀錄")
    mask = (df['Collected'] == 'Y')
    df_show = df[mask].copy()
    
    # === 套用批量選取邏輯 ===
    df_show = process_batch_selection(df_show, "撤銷", "tab4")
    
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
            st.session_state["select_all_tab4"] = False # 重置全選
            st.rerun()

# ---------------- TAB 5: 不符名單 ----------------
with tab5:
    st.subheader("不符資格名單")
    mask = ((df['反思會'].str.upper() != 'Y') | (df['反思表'].str.upper() != 'Y')) & (df['DocGeneratedDate'] == '')
    df_show = df[mask].copy()
    
    # === 套用批量選取邏輯 ===
    df_show = process_batch_selection(df_show, "放行", "tab5")
    
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
            st.session_state["select_all_tab5"] = False # 重置全選
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
