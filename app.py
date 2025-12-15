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
            if 'ID序號' in df.columns:
                df['ID序號'] = df['ID序號'].astype(str).str.strip()
            else:
                if len(df.columns) > 0:
                    df.rename(columns={df.columns[0]: 'ID序號'}, inplace=True)
                    df['ID序號'] = df['ID序號'].astype(str).str.strip()
            for col in SYSTEM_COLS:
                if col not in df.columns: df[col] = ''
            df = df.fillna('')
        else:
            df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)
        return df
    except:
        return pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

# ================= 主程式 =================
st.title("☁️ 實習津貼管理系統 (V41 資料編輯版)")

# 側邊欄
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
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📥 建立新表", 
    "📄 [1] 準備匯出", 
    "✅ [2] 確認領取", 
    "🚫 [3] 不符名單",
    "🛠️ 進階管理",
    "✏️ 修改資料"  # New Tab
])

# ---------------- Tab 1: 建立新表 ----------------
with tab1:
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
                            gc = get_write_client()
                            sh = gc.open_by_url(SPREADSHEET_URL)
                            new_ws = sh.add_worksheet(title=new_sheet_name, rows=len(new_df)+50, cols=20)
                            new_ws.update([new_df.columns.tolist()] + new_df.values.tolist())
                            st.success(f"成功建立「{new_sheet_name}」！")
                            time.sleep(2)
                            st.cache_data.clear()
                            st.rerun()
                    else: st.error("欄位不足")
                except Exception as e: st.error(f"錯誤: {e}")
        else: st.error("請填寫名稱並選擇檔案")

# ---------------- Tab 2: 準備匯出 ----------------
with tab2:
    st.subheader(f"📄 準備匯出 ({selected_sheet_name})")
    if '反思會' in df.columns:
        mask_ready = ((df['反思會'].astype(str).str.strip().str.upper() == 'Y') & 
                      (df['反思表'].astype(str).str.strip().str.upper() == 'Y') & 
                      (df['DocGeneratedDate'] == ''))
        df_ready = df[mask_ready].copy()
        
        df_ready.insert(0, "選取", False)
        edited_ready = st.data_editor(df_ready, column_config={"選取": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df.columns if c!="選取"], hide_index=True, key="ed_ready")
        
        if st.button("📤 匯出 & 更新狀態", type="primary"):
            sel = edited_ready[edited_ready["選取"]==True]
            if not sel.empty:
                try:
                    gc = get_write_client()
                    worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                    today = datetime.now().strftime("%Y-%m-%d")
                    head = worksheet.row_values(1)
                    c_doc = head.index('DocGeneratedDate')+1
                    c_staff = head.index('ResponsibleStaff')+1
                    cloud_ids = [str(x).strip() for x in worksheet.col_values(1)]
                    
                    prog = st.progress(0)
                    ex_list = []
                    for i, (idx, row) in enumerate(sel.iterrows()):
                        tid = str(row['ID序號']).strip()
                        if tid in cloud_ids:
                            row_num = cloud_ids.index(tid) + 1
                            worksheet.update_cell(row_num, c_doc, today)
                            worksheet.update_cell(row_num, c_staff, staff_name)
                            rec = row.to_dict(); del rec['選取']; rec.update({'StaffName':staff_name, 'TodayDate':today})
                            ex_list.append(rec)
                        prog.progress((i+1)/len(sel))
                    
                    if ex_list:
                        out = io.BytesIO()
                        pd.DataFrame(ex_list).to_excel(out, index=False)
                        st.download_button("📥 下載 MailMerge Source", out.getvalue(), "MailMerge_Source.xlsx")
                        st.success("完成！")
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()
                except Exception as e: st.error(f"錯誤: {e}")

# ---------------- Tab 3: 確認領取 ----------------
with tab3:
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
                        tid = str(row['ID序號']).strip()
                        if tid in cloud_ids:
                            row_num = cloud_ids.index(tid) + 1
                            worksheet.update_cell(row_num, c_col, 'Y')
                            worksheet.update_cell(row_num, c_date, now)
                        prog.progress((i+1)/len(sel))
                    st.success("更新完成！")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e: st.error(f"錯誤: {e}")
        
        if st.button("↩️ 退回至準備匯出 (清除日期)"):
            sel = ed_conf[ed_conf["確認"]==True]
            if not sel.empty:
                try:
                    gc = get_write_client()
                    worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                    head = worksheet.row_values(1)
                    c_doc = head.index('DocGeneratedDate')+1
                    c_staff = head.index('ResponsibleStaff')+1
                    cloud_ids = [str(x).strip() for x in worksheet.col_values(1)]
                    
                    for idx, row in sel.iterrows():
                        tid = str(row['ID序號']).strip()
                        if tid in cloud_ids:
                            row_num = cloud_ids.index(tid) + 1
                            worksheet.update_cell(row_num, c_doc, "")
                            worksheet.update_cell(row_num, c_staff, "")
                    st.success("已退回")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e: st.error(f"錯誤: {e}")

# ---------------- Tab 4: 不符名單 ----------------
with tab4:
    st.subheader(f"🚫 不符合資格名單 ({selected_sheet_name})")
    
    if '反思會' in df.columns:
        mask_fail = (((df['反思會'].astype(str).str.strip().str.upper() != 'Y') | 
                      (df['反思表'].astype(str).str.strip().str.upper() != 'Y')) &
                     (df['DocGeneratedDate'] == ''))
        df_fail = df[mask_fail].copy()
        
        if df_fail.empty:
            st.info("太棒了！沒有不符合資格的人員。")
        else:
            st.warning(f"共有 {len(df_fail)} 人條件未達標。")
            df_fail.insert(0, "選取", False)
            ed_fail = st.data_editor(df_fail, column_config={"選取": st.column_config.CheckboxColumn(required=True, label="強制放行")}, disabled=[c for c in df.columns if c != "選取"], hide_index=True, key="ed_fail")
            
            if st.button("➡️ 強制改為合格 (Y/Y) 並移至匯出區", type="primary"):
                sel = ed_fail[ed_fail["選取"]==True]
                if not sel.empty:
                    if st.checkbox("確定要強制修改 Google Sheet 資料為 Y 嗎？"):
                        try:
                            gc = get_write_client()
                            worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                            head = worksheet.row_values(1)
                            c1_idx = head.index('反思會')+1 if '反思會' in head else 7
                            c2_idx = head.index('反思表')+1 if '反思表' in head else 8
                            cloud_ids = [str(x).strip() for x in worksheet.col_values(1)]
                            
                            prog = st.progress(0)
                            for i, (idx, row) in enumerate(sel.iterrows()):
                                tid = str(row['ID序號']).strip()
                                if tid in cloud_ids:
                                    r_num = cloud_ids.index(tid) + 1
                                    worksheet.update_cell(r_num, c1_idx, 'Y')
                                    worksheet.update_cell(r_num, c2_idx, 'Y')
                                prog.progress((i+1)/len(sel))
                            
                            st.success(f"已放行 {len(sel)} 人！請至 Tab 2 進行匯出。")
                            time.sleep(1)
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e: st.error(f"錯誤: {e}")

# ---------------- Tab 5: 進階管理 ----------------
with tab5:
    st.subheader(f"🛠️ 進階管理 - {selected_sheet_name}")
    st.error("⚠️ 危險區域")
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
            if st.button("🔴 是，確認刪除", key="btn_del_sheet"):
                with st.spinner("刪除中..."):
                    gc = get_write_client()
                    sh = gc.open_by_url(SPREADSHEET_URL)
                    worksheet = sh.worksheet(selected_sheet_name)
                    sh.del_worksheet(worksheet)
                    st.success("已刪除！")
                    st.session_state.confirm_del_sheet = False
                    time.sleep(2)
                    st.cache_data.clear()
                    st.rerun()
        with col_no:
            if st.button("取消", key="btn_cancel_sheet"):
                st.session_state.confirm_del_sheet = False
                st.rerun()

# ---------------- Tab 6: 修改資料 (New) ----------------
with tab6:
    st.subheader("✏️ 修改參加者資料")
    st.info("搜尋參加者 -> 點選目標 -> 修改欄位 -> 儲存")
    
    # 1. 搜尋區域
    search_q = st.text_input("🔍 請輸入姓名或 ID：", placeholder="例如: 陳大文 或 102")
    
    selected_person = None
    
    if search_q:
        # 模糊搜尋
        mask_search = (
            df['ID序號'].astype(str).str.contains(search_q, case=False) |
            df['姓名(中文)'].astype(str).str.contains(search_q, case=False) |
            df['姓名(英文)'].astype(str).str.contains(search_q, case=False)
        )
        search_results = df[mask_search]
        
        if search_results.empty:
            st.warning("找不到符合的資料")
        else:
            # 2. 選擇參加者
            person_options = [f"{row['ID序號']} - {row['姓名(中文)']}" for idx, row in search_results.iterrows()]
            selected_option = st.selectbox("👇 請選擇要修改的對象：", person_options)
            
            if selected_option:
                target_id = selected_option.split(" - ")[0]
                person_data = df[df['ID序號'] == target_id].iloc[0]
                
                st.divider()
                st.markdown(f"### 📝 編輯：{person_data['姓名(中文)']}")
                
                # 3. 編輯表單
                with st.form("edit_form"):
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        new_name_chi = st.text_input("姓名 (中文)", value=person_data['姓名(中文)'])
                        new_phone = st.text_input("電話", value=person_data['電話'])
                        new_days = st.text_input("實習日數", value=person_data['實習日數'])
                    with col_e2:
                        new_name_eng = st.text_input("姓名 (英文)", value=person_data['姓名(英文)'])
                        new_cond1 = st.selectbox("反思會 (Y/N)", ["Y", "N", ""], index=["Y", "N", ""].index(person_data['反思會']) if person_data['反思會'] in ["Y", "N", ""] else 2)
                        new_cond2 = st.selectbox("反思表 (Y/N)", ["Y", "N", ""], index=["Y", "N", ""].index(person_data['反思表']) if person_data['反思表'] in ["Y", "N", ""] else 2)
                    
                    submitted = st.form_submit_button("💾 儲存修改")
                    
                    if submitted:
                        try:
                            with st.spinner("正在寫入 Google Sheets..."):
                                gc = get_write_client()
                                worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                                cloud_ids = [str(x).strip() for x in worksheet.col_values(1)]
                                
                                if target_id in cloud_ids:
                                    row_idx = cloud_ids.index(target_id) + 1
                                    
                                    # 取得欄位位置 (Index)
                                    header = worksheet.row_values(1)
                                    
                                    # 建立要更新的 Mapping (欄位名: 新值)
                                    updates = {
                                        '姓名(中文)': new_name_chi,
                                        '姓名(英文)': new_name_eng,
                                        '電話': new_phone,
                                        '實習日數': new_days,
                                        '反思會': new_cond1,
                                        '反思表': new_cond2
                                    }
                                    
                                    for col_name, new_val in updates.items():
                                        if col_name in header:
                                            c_idx = header.index(col_name) + 1
                                            worksheet.update_cell(row_idx, c_idx, new_val)
                                    
                                    st.success(f"✅ {new_name_chi} 的資料已更新！")
                                    time.sleep(1)
                                    st.cache_data.clear() # 清除快取以顯示最新
                                    st.rerun()
                                else:
                                    st.error("雲端找不到此 ID，可能已被刪除。")
                        except Exception as e:
                            st.error(f"更新失敗: {e}")
