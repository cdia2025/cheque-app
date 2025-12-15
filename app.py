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

# ================= 初始化 Session State =================
if 'df_main' not in st.session_state: st.session_state.df_main = None
if 'current_sheet' not in st.session_state: st.session_state.current_sheet = None

# 控制刪除確認框
if 'show_del_confirm' not in st.session_state: st.session_state.show_del_confirm = False
if 'show_clear_confirm' not in st.session_state: st.session_state.show_clear_confirm = False
if 'show_sheet_confirm' not in st.session_state: st.session_state.show_sheet_confirm = False

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

# ================= 核心函式 =================
def fetch_data_from_cloud(sheet_name):
    try:
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl=0)
        if not df.empty:
            # 1. 清理欄位名稱空白
            df.columns = df.columns.str.strip()
            
            # 2. ID 強制轉字串 (關鍵!)
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
st.title("☁️ 實習津貼管理系統 (V38 強力刪除版)")

# --- 側邊欄 ---
with st.sidebar:
    st.header("🎛️ 設定面板")
    staff_name = st.text_input("👤 負責職員姓名 (必填)", key="staff_input")
    st.divider()
    
    try:
        gc = get_write_client()
        sh = gc.open_by_url(SPREADSHEET_URL)
        sheet_names = [ws.title for ws in sh.worksheets()]
        
        idx = 0
        if st.session_state.current_sheet in sheet_names:
            idx = sheet_names.index(st.session_state.current_sheet)
        elif len(sheet_names) > 0:
            idx = 0
            
        selected_sheet_name = st.selectbox("📂 選擇工作表", sheet_names, index=idx)
    except Exception as e:
        st.error(f"連線失敗: {e}")
        st.stop()

    if st.button("🔄 重新整理資料"):
        st.cache_data.clear()
        st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
        st.session_state.current_sheet = selected_sheet_name
        # 重置確認狀態
        st.session_state.show_del_confirm = False
        st.session_state.show_clear_confirm = False
        st.session_state.show_sheet_confirm = False
        st.rerun()

    # 自動載入
    if st.session_state.df_main is None or st.session_state.current_sheet != selected_sheet_name:
        with st.spinner(f"讀取中..."):
            st.session_state.df_main = fetch_data_from_cloud(selected_sheet_name)
            st.session_state.current_sheet = selected_sheet_name

if not staff_name:
    st.warning("⚠️ 請先在左側輸入您的姓名。")
    st.stop()

df = st.session_state.df_main

try:
    worksheet = sh.worksheet(selected_sheet_name)
except:
    st.warning("工作表讀取中，請稍候或重新整理...")
    st.cache_data.clear()
    time.sleep(1)
    st.rerun()

# ================= 分頁 =================
tab_upload, tab_prepare, tab_confirm, tab_manage = st.tabs([
    "📥 建立新表", "📄 [1] 匯出", "✅ [2] 領取", "🛠️ 刪除管理"
])

# ---------------- Tab 1: 建立新表 ----------------
with tab_upload:
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
                            new_ws = sh.add_worksheet(title=new_sheet_name, rows=len(new_df)+50, cols=20)
                            new_ws.update([new_df.columns.tolist()] + new_df.values.tolist())
                            st.success(f"成功建立「{new_sheet_name}」！")
                            time.sleep(1)
                            st.cache_data.clear()
                            st.rerun()
                    else: st.error("欄位不足")
                except Exception as e: st.error(f"錯誤: {e}")
        else: st.error("請填寫名稱並選擇檔案")

# ---------------- Tab 2: 準備匯出 ----------------
with tab_prepare:
    st.subheader(f"📄 準備匯出 ({selected_sheet_name})")
    if '反思會' in df.columns:
        mask_ready = ((df['反思會'].astype(str).str.upper() == 'Y') & (df['反思表'].astype(str).str.upper() == 'Y') & (df['DocGeneratedDate'] == ''))
        df_ready = df[mask_ready].copy()
        df_ready.insert(0, "選取", False)
        edited_ready = st.data_editor(df_ready, column_config={"選取": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df.columns if c!="選取"], hide_index=True, key="ed_ready")
        
        if st.button("📤 匯出 & 更新狀態", type="primary"):
            sel = edited_ready[edited_ready["選取"]==True]
            if not sel.empty:
                today = datetime.now().strftime("%Y-%m-%d")
                head = worksheet.row_values(1)
                try:
                    c_doc = head.index('DocGeneratedDate')+1
                    c_staff = head.index('ResponsibleStaff')+1
                    prog = st.progress(0)
                    ex_list = []
                    
                    # 抓取雲端所有 ID 以便快速比對
                    cloud_ids = [str(x).strip() for x in worksheet.col_values(1)] 
                    
                    for i, (idx, row) in enumerate(sel.iterrows()):
                        target_id = str(row['ID序號']).strip()
                        try:
                            # 比對 ID (注意：cloud_ids index 0 是 row 1)
                            if target_id in cloud_ids:
                                row_num = cloud_ids.index(target_id) + 1
                                worksheet.update_cell(row_num, c_doc, today)
                                worksheet.update_cell(row_num, c_staff, staff_name)
                                
                                rec = row.to_dict(); del rec['選取']; rec.update({'StaffName':staff_name, 'TodayDate':today})
                                ex_list.append(rec)
                        except: pass
                        prog.progress((i+1)/len(sel))
                    
                    if ex_list:
                        out = io.BytesIO()
                        pd.DataFrame(ex_list).to_excel(out, index=False)
                        st.download_button("📥 下載 MailMerge Source", out.getvalue(), "MailMerge_Source.xlsx")
                        st.success("完成！")
                        time.sleep(1)
                        st.rerun()
                except: st.error("雲端欄位對應錯誤")

# ---------------- Tab 3: 確認領取 ----------------
with tab_confirm:
    st.subheader(f"✅ 確認領取 ({selected_sheet_name})")
    if 'Collected' in df.columns:
        mask_conf = ((df['DocGeneratedDate']!='') & (df['Collected']!='Y'))
        df_conf = df[mask_conf].copy()
        df_conf.insert(0, "確認", False)
        ed_conf = st.data_editor(df_conf, column_config={"確認": st.column_config.CheckboxColumn(required=True)}, disabled=[c for c in df.columns if c!="確認"], hide_index=True, key="ed_conf")
        
        if st.button("✅ 確認已取票", type="primary"):
            sel = ed_conf[ed_conf["確認"]==True]
            if not sel.empty:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                head = worksheet.row_values(1)
                c_col = head.index('Collected')+1
                c_date = head.index('CollectedDate')+1
                
                # 抓取雲端所有 ID
                cloud_ids = [str(x).strip() for x in worksheet.col_values(1)]
                
                prog = st.progress(0)
                for i, (idx, row) in enumerate(sel.iterrows()):
                    target_id = str(row['ID序號']).strip()
                    try:
                        if target_id in cloud_ids:
                            row_num = cloud_ids.index(target_id) + 1
                            worksheet.update_cell(row_num, c_col, 'Y')
                            worksheet.update_cell(row_num, c_date, now)
                    except: pass
                    prog.progress((i+1)/len(sel))
                st.success("更新完成！")
                st.rerun()

# ---------------- Tab 4: 刪除管理 (強力修復版) ----------------
with tab_manage:
    st.subheader(f"🛠️ 資料管理 - {selected_sheet_name}")
    st.error("⚠️ 危險操作區：請謹慎使用")
    
    df_del = df.copy()
    df_del.insert(0, "刪除", False)
    ed_del = st.data_editor(df_del, column_config={"刪除": st.column_config.CheckboxColumn(required=True, label="選取")}, hide_index=True, key="ed_del")
    
    st.divider()
    c1, c2, c3 = st.columns(3)
    
    # === 功能 1: 刪除選取列 (修復版) ===
    with c1:
        st.markdown("##### 🗑️ 刪除選取列")
        if st.button("請求刪除選取資料"):
            sel_rows = ed_del[ed_del["刪除"]==True]
            if sel_rows.empty:
                st.toast("請先勾選上方的資料！", icon="⚠️")
            else:
                st.session_state.show_del_confirm = True
                st.rerun()
        
        if st.session_state.show_del_confirm:
            st.warning("確定要刪除勾選資料？(無法復原)")
            col_yes, col_no = st.columns(2)
            with col_yes:
                if st.button("🔴 確認刪除", key="btn_confirm_del"):
                    with st.spinner("正在刪除...請勿關閉視窗"):
                        sel_rows = ed_del[ed_del["刪除"]==True]
                        
                        # 1. 抓取雲端所有 ID (第一欄)
                        # 這比 cell.find 快且準確
                        cloud_ids = [str(x).strip() for x in worksheet.col_values(1)]
                        
                        rows_to_del = []
                        for idx, row in sel_rows.iterrows():
                            target_id = str(row['ID序號']).strip()
                            # 找出這個 ID 在雲端的所有位置 (可能有重複，都刪)
                            # enumerate index 從 0 開始，sheet row 從 1 開始 -> index + 1
                            matched_rows = [i + 1 for i, x in enumerate(cloud_ids) if x == target_id]
                            rows_to_del.extend(matched_rows)
                        
                        # 去重並排序 (由大到小)
                        rows_to_del = sorted(list(set(rows_to_del)), reverse=True)
                        
                        # 執行刪除
                        for r in rows_to_del:
                            worksheet.delete_rows(r)
                        
                        st.success(f"已刪除 {len(rows_to_del)} 筆資料")
                        st.session_state.show_del_confirm = False
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()
            with col_no:
                if st.button("取消", key="btn_cancel_del"):
                    st.session_state.show_del_confirm = False
                    st.rerun()

    # === 功能 2: 清空整表 (修復版) ===
    with c2:
        st.markdown("##### 🧹 清空內容")
        if st.button("請求清空內容"):
            st.session_state.show_clear_confirm = True
            st.rerun()
            
        if st.session_state.show_clear_confirm:
            st.warning("確定清空？只會保留標題列。")
            col_yes2, col_no2 = st.columns(2)
            with col_yes2:
                if st.button("🔴 確認清空", key="btn_confirm_clear"):
                    with st.spinner("清空中..."):
                        # 1. 備份標題 (從目前的 DF 拿)
                        headers = df.columns.tolist()
                        # 2. 清空所有
                        worksheet.clear()
                        # 3. 寫回標題
                        worksheet.append_row(headers)
                        
                        st.success("已清空內容！")
                        st.session_state.show_clear_confirm = False
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()
            with col_no2:
                if st.button("取消", key="btn_cancel_clear"):
                    st.session_state.show_clear_confirm = False
                    st.rerun()

    # === 功能 3: 刪除工作表 (維持原樣) ===
    with c3:
        st.markdown("##### 🔥 刪除本工作表")
        if st.button("請求刪除工作表"):
            if len(sheet_names) <= 1:
                st.error("這是最後一張表，無法刪除。")
            else:
                st.session_state.show_sheet_confirm = True
                st.rerun()
        
        if st.session_state.show_sheet_confirm:
            st.warning(f"確定永久刪除「{selected_sheet_name}」？")
            col_yes3, col_no3 = st.columns(2)
            with col_yes3:
                if st.button("🔴 確認刪除", key="btn_confirm_sheet"):
                    with st.spinner("刪除中..."):
                        sh.del_worksheet(worksheet)
                        st.success("工作表已刪除")
                        st.session_state.show_sheet_confirm = False
                        st.session_state.current_sheet = None
                        time.sleep(1)
                        st.cache_data.clear()
                        st.rerun()
            with col_no3:
                if st.button("取消", key="btn_cancel_sheet"):
                    st.session_state.show_sheet_confirm = False
                    st.rerun()
