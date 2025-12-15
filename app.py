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

# ================= 核心：萬能 ID 清洗函式 =================
def clean_id(val):
    if val is None: return ""
    s = str(val).strip()
    if s == "": return ""
    if s.endswith(".0"): return s[:-2]
    return s

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

# ================= 快取函式 =================
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
            if 'ID序號' not in df.columns and len(df.columns) > 0:
                df.rename(columns={df.columns[0]: 'ID序號'}, inplace=True)
            
            if 'ID序號' in df.columns:
                df['ID序號'] = df['ID序號'].apply(clean_id)
            
            for col in SYSTEM_COLS:
                if col not in df.columns: df[col] = ''
            df = df.fillna('')
        else:
            df = pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)
        return df
    except:
        return pd.DataFrame(columns=REQUIRED_COLS + SYSTEM_COLS)

# ================= 統計計算 =================
def calculate_stats(df):
    if df.empty or '反思會' not in df.columns:
        return {'總人數': 0, '待匯出': 0, '待領取': 0, '已完成': 0, '不符資格': 0}
    c1 = df['反思會'].astype(str).str.strip().str.upper()
    c2 = df['反思表'].astype(str).str.strip().str.upper()
    doc = df['DocGeneratedDate'].astype(str).str.strip()
    done = df['Collected'].astype(str).str.strip().str.upper()
    is_eligible = (c1 == 'Y') & (c2 == 'Y')
    return {
        '總人數': len(df),
        '待匯出': (is_eligible & (doc == '')).sum(),
        '待領取': ((doc != '') & (done != 'Y')).sum(),
        '已完成': (done == 'Y').sum(),
        '不符資格': ((~is_eligible) & (doc == '')).sum()
    }

# ================= 主程式 =================
st.title("☁️ 實習津貼管理系統 (V48 修復版)")

# Session State 初始化
if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None
if 'export_success_file' not in st.session_state: st.session_state.export_success_file = None

# --- 側邊欄 ---
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
    
    if st.session_state.last_selected_sheet != selected_sheet_name:
        st.session_state.export_success_file = None
        st.session_state.last_selected_sheet = selected_sheet_name

    if st.button("🔄 重新整理資料"):
        st.cache_data.clear()
        st.session_state.export_success_file = None
        st.rerun()

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名。")
    st.stop()

df = fetch_data_cached(selected_sheet_name)

# ================= 分頁 =================
# 這裡不使用 st.tabs 的回傳值，因為我們要手動控制顯示內容
# 但 Streamlit 預設沒有 programmatic tab selection，所以我們用標準 tabs
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📥 建立新表", 
    "📄 [1] 準備匯出", 
    "🔵 [2] 待領取", 
    "🟢 [3] 已取票清單", 
    "🚫 [4] 不符", 
    "🛠️ 管理", 
    "✏️ 修改", 
    "📊 統計"
])

# ---------------- Tab 1: 建立新表 ----------------
with tab1:
    st.subheader("📥 上傳 Excel")
    uploaded_file = st.file_uploader("選擇 Excel 檔案", type=['xlsx', 'xls'], key="upl")
    new_sheet_name = st.text_input("輸入新工作表名稱", placeholder="2024_第一期", key="new_s_in")
    
    if st.button("🚀 建立並上傳", type="primary"):
        if uploaded_file and new_sheet_name:
            if new_sheet_name in sheet_names: st.error("名稱已存在！")
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
                        new_df['ID序號'] = new_df['ID序號'].apply(clean_id)
                        new_df = new_df.fillna('')
                        
                        with st.spinner("建立中..."):
                            gc = get_write_client()
                            sh = gc.open_by_url(SPREADSHEET_URL)
                            new_ws = sh.add_worksheet(title=new_sheet_name, rows=len(new_df)+50, cols=20)
                            data_to_write = [new_df.columns.tolist()] + new_df.astype(str).values.tolist()
                            new_ws.update(data_to_write)
                            st.success(f"成功建立「{new_sheet_name}」！")
                            # 更新選擇狀態，避免跳頁
                            st.session_state.last_selected_sheet = new_sheet_name
                            time.sleep(2); st.cache_data.clear(); st.rerun()
                    else: st.error("欄位不足")
                except Exception as e: st.error(f"錯誤: {e}")
        else: st.error("請填寫名稱並選擇檔案")

# ---------------- Tab 2: 準備匯出 ----------------
with tab2:
    st.subheader(f"📄 準備匯出 ({selected_sheet_name})")
    
    if st.session_state.export_success_file:
        st.success("✅ 匯出成功！請下載：")
        st.download_button("📥 下載 MailMerge_Source.xlsx", st.session_state.export_success_file, "MailMerge_Source.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        st.divider()

    if '反思會' in df.columns:
        mask_ready = ((df['反思會'].astype(str).str.strip().str.upper() == 'Y') & 
                      (df['反思表'].astype(str).str.strip().str.upper() == 'Y') & 
                      (df['DocGeneratedDate'] == ''))
        df_ready = df[mask_ready].copy()
        df_ready.insert(0, "選取", False)
        
        edited_ready = st.data_editor(df_ready, column_config={"選取": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df.columns if c!="選取"], hide_index=True, key="ed_ready")
        
        if st.button("📤 匯出資料 & 更新雲端"):
            sel = edited_ready[edited_ready["選取"]==True]
            if sel.empty: st.warning("❌ 未選取")
            else:
                try:
                    with st.spinner("更新中..."):
                        gc = get_write_client(); worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                        today = datetime.now().strftime("%Y-%m-%d"); head = worksheet.row_values(1)
                        if 'DocGeneratedDate' not in head: st.error("缺欄位"); st.stop()
                        
                        # 動態找 ID 欄位 (V48 修復)
                        try:
                            id_col_idx = head.index('ID序號') + 1
                        except:
                            id_col_idx = 1 # 預設第一欄
                        
                        c_doc = head.index('DocGeneratedDate')+1; c_staff = head.index('ResponsibleStaff')+1
                        raw_ids = worksheet.col_values(id_col_idx); cloud_ids = [clean_id(x) for x in raw_ids]
                        prog = st.progress(0); ex_list = []
                        
                        for i, (idx, row) in enumerate(sel.iterrows()):
                            tid = clean_id(row['ID序號'])
                            if tid in cloud_ids:
                                r_num = cloud_ids.index(tid) + 1
                                worksheet.update_cell(r_num, c_doc, today)
                                worksheet.update_cell(r_num, c_staff, staff_name)
                                rec = row.to_dict(); del rec['選取']; rec.update({'StaffName':staff_name, 'TodayDate':today}); ex_list.append(rec)
                            prog.progress((i+1)/len(sel))
                        
                    if ex_list:
                        out = io.BytesIO(); pd.DataFrame(ex_list).to_excel(out, index=False)
                        st.session_state.export_success_file = out.getvalue()
                        st.toast("匯出成功！")
                        time.sleep(1); st.cache_data.clear(); st.rerun()
                    else: st.error("找不到 ID")
                except Exception as e: st.error(f"錯誤: {e}")

# ---------------- Tab 3: 待領取 ----------------
with tab3:
    st.subheader(f"🔵 待領取支票名單 ({selected_sheet_name})")
    
    if 'Collected' in df.columns:
        mask_conf = ((df['DocGeneratedDate']!='') & (df['Collected']!='Y'))
        df_conf = df[mask_conf].copy()
        df_conf.insert(0, "確認", False)
        ed_conf = st.data_editor(df_conf, column_config={"確認": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df.columns if c!="確認"], hide_index=True, key="ed_conf")
        
        c1, c2 = st.columns([1, 2])
        with c1:
            if st.button("✅ 確認已取票", type="primary"):
                sel = ed_conf[ed_conf["確認"]==True]
                if not sel.empty:
                    try:
                        with st.spinner("更新中..."):
                            gc = get_write_client(); worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); head = worksheet.row_values(1)
                            
                            # V48: 動態找欄位
                            try: id_col_idx = head.index('ID序號') + 1
                            except: id_col_idx = 1
                            
                            c_col = head.index('Collected')+1; c_date = head.index('CollectedDate')+1
                            cloud_ids = [clean_id(x) for x in worksheet.col_values(id_col_idx)]
                            prog = st.progress(0)
                            
                            for i, (idx, row) in enumerate(sel.iterrows()):
                                tid = clean_id(row['ID序號'])
                                if tid in cloud_ids:
                                    r_num = cloud_ids.index(tid)+1
                                    worksheet.update_cell(r_num, c_col, 'Y')
                                    worksheet.update_cell(r_num, c_date, now)
                                prog.progress((i+1)/len(sel))
                            st.success("完成！"); time.sleep(1); st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(f"錯誤: {e}")
        
        with c2:
            if st.button("↩️ 退回至準備匯出"):
                sel = ed_conf[ed_conf["確認"]==True]
                if not sel.empty:
                    if st.checkbox("確定退回？(清除日期)"):
                        try:
                            with st.spinner("退回中..."):
                                gc = get_write_client(); worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                                head = worksheet.row_values(1)
                                
                                try: id_col_idx = head.index('ID序號') + 1
                                except: id_col_idx = 1
                                
                                c_doc = head.index('DocGeneratedDate')+1; c_staff = head.index('ResponsibleStaff')+1
                                cloud_ids = [clean_id(x) for x in worksheet.col_values(id_col_idx)]
                                
                                for i, (idx, row) in enumerate(sel.iterrows()):
                                    tid = clean_id(row['ID序號'])
                                    if tid in cloud_ids:
                                        r = cloud_ids.index(tid) + 1
                                        worksheet.update_cell(r, c_doc, "")
                                        worksheet.update_cell(r, c_staff, "")
                                st.success("已退回"); time.sleep(1); st.cache_data.clear(); st.rerun()
                        except Exception as e: st.error(f"錯誤: {e}")

# ---------------- Tab 4: 已取票清單 ----------------
with tab4:
    st.subheader(f"🟢 已取票清單 ({selected_sheet_name})")
    
    if 'Collected' in df.columns:
        mask_done = (df['Collected'] == 'Y')
        df_done = df[mask_done].copy()
        
        if df_done.empty:
            st.info("無紀錄")
        else:
            df_done.insert(0, "選取", False)
            ed_done = st.data_editor(df_done, column_config={"選取": st.column_config.CheckboxColumn(required=True, label="撤銷")}, disabled=[c for c in df.columns if c!="選取"], hide_index=True, key="ed_done")
            
            if st.button("↩️ 撤銷領取 (退回 Tab 2)"):
                sel = ed_done[ed_done["選取"]==True]
                if not sel.empty:
                    if st.checkbox("確定撤銷？"):
                        try:
                            with st.spinner("撤銷中..."):
                                gc = get_write_client(); worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                                head = worksheet.row_values(1)
                                
                                try: id_col_idx = head.index('ID序號') + 1
                                except: id_col_idx = 1
                                
                                c_col = head.index('Collected')+1; c_date = head.index('CollectedDate')+1
                                cloud_ids = [clean_id(x) for x in worksheet.col_values(id_col_idx)]
                                
                                for i, (idx, row) in enumerate(sel.iterrows()):
                                    tid = clean_id(row['ID序號'])
                                    if tid in cloud_ids:
                                        r = cloud_ids.index(tid) + 1
                                        worksheet.update_cell(r, c_col, "")
                                        worksheet.update_cell(r, c_date, "")
                                st.success("已撤銷"); time.sleep(1); st.cache_data.clear(); st.rerun()
                        except Exception as e: st.error(f"錯誤: {e}")

# ---------------- Tab 5: 不符名單 ----------------
with tab5:
    st.subheader(f"🚫 不符資格 ({selected_sheet_name})")
    if '反思會' in df.columns:
        mask_fail = (((df['反思會'].astype(str).str.strip().str.upper() != 'Y') | (df['反思表'].astype(str).str.strip().str.upper() != 'Y')) & (df['DocGeneratedDate'] == ''))
        df_fail = df[mask_fail].copy()
        if df_fail.empty: st.info("無不符資格人員")
        else:
            df_fail.insert(0, "選取", False)
            ed_fail = st.data_editor(df_fail, column_config={"選取": st.column_config.CheckboxColumn(required=True, label="放行")}, disabled=[c for c in df.columns if c != "選取"], hide_index=True, key="ed_fail")
            if st.button("➡️ 強制放行 (改Y)", type="primary"):
                sel = ed_fail[ed_fail["選取"]==True]
                if not sel.empty:
                    if st.checkbox("確認修改？"):
                        try:
                            gc = get_write_client(); worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                            head = worksheet.row_values(1)
                            
                            try: id_col_idx = head.index('ID序號') + 1
                            except: id_col_idx = 1
                            
                            c1 = head.index('反思會')+1 if '反思會' in head else 7; c2 = head.index('反思表')+1 if '反思表' in head else 8
                            cloud_ids = [clean_id(x) for x in worksheet.col_values(id_col_idx)]; prog = st.progress(0)
                            for i, (idx, row) in enumerate(sel.iterrows()):
                                tid = clean_id(row['ID序號'])
                                if tid in cloud_ids:
                                    r = cloud_ids.index(tid)+1; worksheet.update_cell(r, c1, 'Y'); worksheet.update_cell(r, c2, 'Y')
                                prog.progress((i+1)/len(sel))
                            st.success("已放行"); time.sleep(1); st.cache_data.clear(); st.rerun()
                        except Exception as e: st.error(f"錯誤: {e}")

# ---------------- Tab 6: 進階管理 ----------------
with tab6:
    st.subheader(f"🛠️ 進階管理 - {selected_sheet_name}")
    st.error("⚠️ 危險區域")
    if st.button("🔥 刪除本工作表"):
        if len(sheet_names) <= 1: st.error("無法刪除最後一張表")
        else: st.session_state.confirm_del_sheet = True
    if st.session_state.get('confirm_del_sheet', False):
        st.warning(f"確定永久刪除「{selected_sheet_name}」？")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔴 確認刪除", key="del_s"):
                with st.spinner("刪除中..."):
                    gc = get_write_client(); sh = gc.open_by_url(SPREADSHEET_URL)
                    sh.del_worksheet(sh.worksheet(selected_sheet_name))
                    st.success("已刪除"); st.session_state.confirm_del_sheet = False; time.sleep(2); st.cache_data.clear(); st.rerun()
        with c2:
            if st.button("取消", key="can_s"): st.session_state.confirm_del_sheet = False; st.rerun()

# ---------------- Tab 7: 修改資料 ----------------
with tab7:
    st.subheader(f"✏️ 直接編輯 - {selected_sheet_name}")
    st.info("直接修改，完成後按「儲存」。")
    df_edit = df.copy()
    disabled_cols = ['ID序號', 'Collected', 'DocGeneratedDate', 'CollectedDate', 'ResponsibleStaff']
    edited_df = st.data_editor(
        df_edit,
        column_config={
            "反思會": st.column_config.SelectboxColumn("反思會", options=["Y", "N", ""], required=True),
            "反思表": st.column_config.SelectboxColumn("反思表", options=["Y", "N", ""], required=True),
            "實習日數": st.column_config.NumberColumn("實習日數", min_value=0, max_value=365, step=1),
        },
        disabled=disabled_cols, hide_index=True, use_container_width=True, num_rows="fixed", key="data_editor_main"
    )
    if st.button("💾 儲存全部修改", type="primary"):
        try:
            with st.spinner("寫入中..."):
                gc = get_write_client(); ws = gc.open_by_url(SPREADSHEET_URL).worksheet(selected_sheet_name)
                final_df = edited_df.fillna("")
                final_df['ID序號'] = final_df['ID序號'].astype(str)
                data_to_write = [final_df.columns.tolist()] + final_df.astype(str).values.tolist()
                ws.update(data_to_write)
                st.success("已同步！"); time.sleep(1); st.cache_data.clear(); st.rerun()
        except Exception as e: st.error(f"失敗: {e}")

# ---------------- Tab 8: 統計 ----------------
with tab8:
    st.subheader("📊 統計")
    curr_stats = calculate_stats(df)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("總數", curr_stats['總人數'])
    c2.metric("待匯出", curr_stats['待匯出'], delta_color="inverse")
    c3.metric("待領取", curr_stats['待領取'], delta_color="normal")
    c4.metric("已完成", curr_stats['已完成'])
    c5.metric("不符", curr_stats['不符資格'], delta_color="inverse")
    
    st.divider()
    if st.button("🚀 掃描所有工作表"):
        with st.spinner("掃描中..."):
            all_data = []
            for sheet in sheet_names:
                try:
                    time.sleep(0.5)
                    sub_df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet, ttl=600)
                    for c in SYSTEM_COLS: 
                        if c not in sub_df.columns: sub_df[c]=''
                    sub_df = sub_df.fillna('')
                    stats = calculate_stats(sub_df)
                    all_data.append({'工作表': sheet, '🟠 待匯出': stats['待匯出'], '🔵 待領取': stats['待領取'], '🟢 已完成': stats['已完成']})
                except: pass
            st.dataframe(pd.DataFrame(all_data), use_container_width=True)
