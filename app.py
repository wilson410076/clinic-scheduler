import streamlit as st
import pandas as pd
import io
import calendar
import openpyxl
from openpyxl.styles import Font
import datetime
import json  # 🌟 新增：用來處理休假備份檔的工具

# 統一字型設定（與原班表一致）
CELL_FONT = Font(name="微軟正黑體", size=10)

# ==========================================
# 1. 診所排班規則與設定
# ==========================================
VALID_DOCTORS = [
    "楊忠霖", "陳逸陽", "官俊彥", "李昆晏", "劉善總", "劉庭禎", "王荷若", "鄭竣文",
    "李亞凡", "陳明志", "黃菁菁", "傅超俊", "陳昶安", "陳苡瑜", "許雅茹", "蔣宜蓁",
    "劉筑昀", "胡瑋凡"
]

DOCTOR_ASSISTANT_MATCH = {
    "陳逸陽": "映璇", "王荷若": "萃屏", "李昆晏": "萃屏", "許雅茹": "和芸", "劉筑昀": "姿穎",
}
SATURDAY_SPECIAL_MATCH = {
    "官俊彥": "濘安", "陳明志": "菀庭",
}

ASSISTANTS = ["映璇", "和芸", "欣寧", "萃屏", "維珍", "菀庭", "姿穎", "濘安"]
COUNTER_PRIORITY = ["欣寧", "維珍", "和芸", "映璇", "姿穎"]

ONLY_COUNTER   = ["欣寧"]                
NO_COUNTER     = ["萃屏", "菀庭", "濘安"]   
NO_NIGHT_SHIFT = ["維珍"]                
MAX_SHIFTS = 42  

st.set_page_config(page_title="恩霖診所 - 自動排班系統", layout="wide")

if "authenticated" not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.title("🔒 恩霖診所 - 排班系統登入")
    pwd = st.text_input("請輸入密碼：", type="password")
    if st.button("登入"):
        if pwd == "115": st.session_state.authenticated = True; st.rerun()
        else: st.error("密碼錯誤！")
    st.stop()

if "timeoff_db" not in st.session_state: st.session_state.timeoff_db = {}
if "shift_stats" not in st.session_state: st.session_state.shift_stats = {}

st.title("🏥 恩霖診所 - 自動排班系統 V18")
st.info("🛡️ 已新增「休假紀錄備份與還原」功能，更新系統不再怕資料遺失！")

tab1, tab2, tab3 = st.tabs(["📁 1. 上傳班表", "📝 2. 助理劃休", "🚀 3. AI 排班"])

with tab1:
    uploaded_file = st.file_uploader("上傳醫師班表 (Excel .xlsx)", type=["xlsx"])
    if uploaded_file:
        try:
            wb_check = openpyxl.load_workbook(uploaded_file, data_only=True)
            ws_check = wb_check.active
            uploaded_file.seek(0)
            found_dates = [cell.value for row in ws_check.iter_rows() for cell in row if isinstance(cell.value, datetime.datetime)]
            if found_dates:
                ref = found_dates[0]
                st.session_state.schedule_year  = ref.year
                st.session_state.schedule_month = ref.month
                sat_days = {d.day for d in found_dates if d.weekday() == 5}
                st.session_state.saturday_dates = sat_days
                st.success(f"✅ 檔案讀取完成！偵測月份：**{ref.year} 年 {ref.month} 月**")
            else: st.warning("⚠️ 找不到日期資料。")
        except Exception as e: st.error(f"讀取失敗：{e}")

with tab2:
    if "saturday_dates" not in st.session_state:
        st.warning("⚠️ 請先上傳班表。")
        st.stop()
    SATURDAY_DATES = st.session_state.saturday_dates
    year, month = st.session_state.schedule_year, st.session_state.schedule_month
    days_in_month = calendar.monthrange(year, month)[1]

    st.info(f"📅 目前班表月份：{year} 年 {month} 月（共 {days_in_month} 天）")

    selected_ast = st.selectbox("選擇助理：", ASSISTANTS)
    if st.session_state.get("last_schedule_month") != (year, month):
        st.session_state.timeoff_db = {}
        st.session_state.last_schedule_month = (year, month)

    if selected_ast not in st.session_state.timeoff_db:
        df = pd.DataFrame({
            "日期": [f"{i}號" for i in range(1, days_in_month + 1)],
            "休整天": [False] * days_in_month, "早休": [False] * days_in_month,
            "午休": [False] * days_in_month, "晚休": [False] * days_in_month,
        })
        if selected_ast in NO_NIGHT_SHIFT: df["晚休"] = True
        st.session_state.timeoff_db[selected_ast] = df

    edited_df = st.data_editor(st.session_state.timeoff_db[selected_ast], hide_index=True, key=f"ed_{selected_ast}_{year}_{month}")
    
    for i, row in edited_df.iterrows():
        day_num = i + 1
        if row["休整天"]:
            edited_df.at[i, "早休"] = True
            edited_df.at[i, "午休"] = True
            if day_num not in SATURDAY_DATES and selected_ast not in NO_NIGHT_SHIFT:
                edited_df.at[i, "晚休"] = True
        if day_num in SATURDAY_DATES: edited_df.at[i, "晚休"] = False

    if st.button(f"💾 儲存 {selected_ast} 休假", type="primary"):
        st.session_state.timeoff_db[selected_ast] = edited_df.copy()
        st.success(f"✅ {selected_ast} 休假已儲存！")

    st.divider()
    
    # 🌟 核心新增：休假備份與還原機制 🌟
    st.subheader("💾 休假紀錄備份與還原")
    st.caption("當系統準備更新或重啟前，請先點擊左側「匯出備份」。更新完畢後，在右側「上傳」該備份檔即可還原所有助理的休假！")
    
    col_export, col_import = st.columns(2)
    
    with col_export:
        if st.session_state.timeoff_db:
            export_data = {}
            for ast, ast_df in st.session_state.timeoff_db.items():
                export_data[ast] = ast_df.to_dict(orient="records")
            json_str = json.dumps(export_data, ensure_ascii=False)
            
            st.download_button(
                label="📥 1. 匯出全部助理休假備份檔",
                data=json_str,
                file_name=f"恩霖診所_休假備份_{year}年{month}月.json",
                mime="application/json",
                use_container_width=True
            )
        else:
            st.button("📥 1. 匯出全部助理休假備份檔", disabled=True, use_container_width=True)
            st.caption("目前尚無資料可匯出")

    with col_import:
        uploaded_backup = st.file_uploader("📤 2. 選擇備份檔還原", type=["json"], label_visibility="collapsed")
        if uploaded_backup is not None:
            if st.button("🔄 執行還原", type="primary", use_container_width=True):
                try:
                    backup_data = json.load(uploaded_backup)
                    for ast, records in backup_data.items():
                        st.session_state.timeoff_db[ast] = pd.DataFrame(records)
                    st.success("✅ 休假紀錄已成功還原！")
                    st.rerun()  # 畫面重整讓資料顯示出來
                except Exception as e:
                    st.error(f"還原失敗，請確認上傳的是正確的 json 備份檔：{e}")

with tab3:
    if st.session_state.shift_stats:
        cols = st.columns(len(ASSISTANTS))
        for idx, ast in enumerate(ASSISTANTS):
            count = st.session_state.shift_stats.get(ast, 0)
            cols[idx].metric(label=ast, value=f"{count} 診")
        st.divider()

    if st.button("🚀 開始智慧排班", type="primary"):
        if not uploaded_file or "saturday_dates" not in st.session_state: st.error("請先上傳班表！")
        else:
            with st.spinner("正在執行精準修正排程..."):
                try:
                    uploaded_file.seek(0)
                    wb = openpyxl.load_workbook(uploaded_file)
                    ws = wb.active
                    rotation_counter = {a: 0 for a in ASSISTANTS}

                    def is_on_leave(ast_name, day_num, shift_type):
                        if ast_name not in st.session_state.timeoff_db or not (1 <= day_num <= 31): return False
                        try: return bool(st.session_state.timeoff_db[ast_name]["早休" if "早" in shift_type else ("午休" if "午" in shift_type else "晚休")].iloc[day_num - 1])
                        except: return False

                    def over_limit(ast_name): return rotation_counter[ast_name] >= MAX_SHIFTS
                    def write_cell(r, c, v): ws.cell(r, c).value = v; ws.cell(r, c).font = CELL_FONT
                    def get_docs(v): return [d for d in VALID_DOCTORS if d in str(v).replace(" ", "")] if v and "休" not in str(v) else []

                    def prevents_split_shift(cand, shift_name, worked_shifts):
                        if shift_name == "晚班":
                            if "早班" in worked_shifts[cand] and "午班" not in worked_shifts[cand]:
                                return False
                        return True

                    def is_continuous(cand, shift_name, worked_shifts):
                        if shift_name == "午班" and "早班" in worked_shifts[cand]: return True
                        if shift_name == "晚班" and "午班" in worked_shifts[cand]: return True
                        return False

                    def ok_for_assist(cand, w_now, day_num, s_name, d_counts, w_shifts):
                        return (cand not in w_now and not is_on_leave(cand, day_num, s_name) and not over_limit(cand)
                                and d_counts.get(cand, 0) < 2 and cand not in ONLY_COUNTER 
                                and not ("晚" in s_name and cand in NO_NIGHT_SHIFT) and prevents_split_shift(cand, s_name, w_shifts))

                    def ok_for_counter(cand, w_now, day_num, s_name, d_counts, w_shifts):
                        return (cand not in w_now and not is_on_leave(cand, day_num, s_name) and not over_limit(cand)
                                and cand not in NO_COUNTER
                                and d_counts.get(cand, 0) < 2 and not ("晚" in s_name and cand in NO_NIGHT_SHIFT) 
                                and prevents_split_shift(cand, s_name, w_shifts))

                    enlin_rows = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value and "恩霖" in str(ws.cell(r, 1).value)]
                    if not enlin_rows: st.error("找不到班表格式"); st.stop()

                    for week_idx, week_start in enumerate(enlin_rows):
                        week_end = enlin_rows[week_idx + 1] if week_idx + 1 < len(enlin_rows) else ws.max_row + 1
                        date_cols = [(c, c+1, v.day, v.weekday()==5) for c in range(2, ws.max_column+1, 2) if isinstance((v:=ws.cell(week_start, c).value), datetime.datetime)]
                        if not date_cols: continue

                        row_labels = {k: r for r in range(week_start, week_end) for k in ["早班","早櫃2","早櫃1","午班","午櫃2","午櫃1","晚班","晚櫃2","晚櫃1","備註"] if k in str(ws.cell(r, 1).value or "").replace(" ", "")}
                        shifts = []
                        for s_name, s_doc, c2, c1 in [("早班","早班","早櫃2","早櫃1"), ("午班","早櫃1","午櫃2","午櫃1"), ("晚班","晚班","晚櫃2","晚櫃1")]:
                            if s_doc in row_labels and c2 in row_labels:
                                shifts.append((s_name, [r for r in range(row_labels[s_doc]+1, row_labels[c2]) if str(ws.cell(r, 1).value or "").strip() in ("11","21","22","23","24")], [x for x in [row_labels.get(c2), row_labels.get(c1)] if x]))

                        for doc_col, asst_col, day_num, is_saturday in date_cols:
                            doc_day_memory = {}
                            daily_shift_counts = {a: 0 for a in ASSISTANTS}
                            daily_worked_shifts = {a: set() for a in ASSISTANTS}

                            for shift_name, doc_rows, counter_rows in shifts:
                                docs_this_shift = [(r, doc) for r in doc_rows for doc in get_docs(ws.cell(r, doc_col).value)]
                                if not docs_this_shift: continue

                                working_now = set()
                                for r in doc_rows + counter_rows:
                                    ea = ws.cell(r, asst_col).value
                                    if ea and str(ea).strip() in ASSISTANTS:
                                        ast = str(ea).strip()
                                        working_now.add(ast)
                                        if shift_name not in daily_worked_shifts[ast]:
                                            daily_shift_counts[ast] += 1
                                            rotation_counter[ast] += 1
                                            daily_worked_shifts[ast].add(shift_name)
                                        if r in doc_rows:
                                            for doc in get_docs(ws.cell(r, doc_col).value): doc_day_memory[doc] = ast

                                for r, doc in docs_this_shift:
                                    if ws.cell(r, asst_col).value: continue 
                                    assigned = next((c for c in [doc_day_memory.get(doc), SATURDAY_SPECIAL_MATCH.get(doc) if is_saturday else None, DOCTOR_ASSISTANT_MATCH.get(doc) if not is_saturday else None] if c and ok_for_assist(c, working_now, day_num, shift_name, daily_shift_counts, daily_worked_shifts)), None)
                                    if not assigned:
                                        pool = sorted([a for a in ASSISTANTS if ok_for_assist(a, working_now, day_num, shift_name, daily_shift_counts, daily_worked_shifts)], 
                                                      key=lambda a: (0 if is_continuous(a, shift_name, daily_worked_shifts) else 1, rotation_counter[a]))
                                        if pool: assigned = pool[0]
                                    if assigned:
                                        write_cell(r, asst_col, assigned); working_now.add(assigned); doc_day_memory[doc] = assigned
                                        rotation_counter[assigned] += 1; daily_shift_counts[assigned] += 1; daily_worked_shifts[assigned].add(shift_name)
                                    else: write_cell(r, asst_col, "缺")

                                target_count = 2 if len(docs_this_shift) >= 4 else 1
                                pre_assigned = sum(1 for cr in counter_rows if ws.cell(cr, asst_col).value and str(ws.cell(cr, asst_col).value).strip() in ASSISTANTS)
                                to_add = target_count - pre_assigned
                                assigned_new = []

                                if to_add > 0:
                                    avail_counters = [c for c in COUNTER_PRIORITY if ok_for_counter(c, working_now, day_num, shift_name, daily_shift_counts, daily_worked_shifts)]
                                    avail_counters.sort(key=lambda a: (0 if is_continuous(a, shift_name, daily_worked_shifts) else 1, rotation_counter.get(a, 0)))
                                    
                                    for cand in avail_counters:
                                        if len(assigned_new) >= to_add: break
                                        assigned_new.append(cand); working_now.add(cand)
                                        rotation_counter[cand] += 1; daily_shift_counts[cand] += 1; daily_worked_shifts[cand].add(shift_name)
                                    
                                    while len(assigned_new) < to_add:
                                        pool = sorted([a for a in ASSISTANTS if ok_for_counter(a, working_now, day_num, shift_name, daily_shift_counts, daily_worked_shifts)], 
                                                      key=lambda a: (0 if is_continuous(a, shift_name, daily_worked_shifts) else 1, rotation_counter[a]))
                                        if pool:
                                            assigned_new.append(pool[0]); working_now.add(pool[0])
                                            rotation_counter[pool[0]] += 1; daily_shift_counts[pool[0]] += 1; daily_worked_shifts[pool[0]].add(shift_name)
                                        else: assigned_new.append("缺"); break

                                c_idx, filled = 0, pre_assigned
                                for cr in counter_rows:
                                    if not ws.cell(cr, asst_col).value or str(ws.cell(cr, asst_col).value).strip() not in ASSISTANTS:
                                        if c_idx < len(assigned_new): write_cell(cr, asst_col, assigned_new[c_idx]); c_idx += 1; filled += 1
                                        elif filled < target_count: write_cell(cr, asst_col, "缺"); filled += 1

                    st.session_state.shift_stats = dict(rotation_counter)
                    s_year, s_month = st.session_state.schedule_year, st.session_state.schedule_month
                    output = io.BytesIO(); wb.save(output)
                    st.success("✅ 排班完成！")
                    st.download_button("📥 下載排班結果", output.getvalue(), f"恩霖診所_{s_year}年{s_month}月排班結果_V18.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e: st.error(f"排班發生錯誤：{e}"); import traceback; st.code(traceback.format_exc())
