import streamlit as st
import pandas as pd
import io
import openpyxl
import datetime
import re

# ==========================================
# 1. 診所排班規則與設定
# ==========================================
# 🩺 醫師專屬白名單 (精準過濾紅框診間號碼與雜訊)
VALID_DOCTORS = [
    "楊忠霖", "陳逸陽", "官俊彥", "李昆晏", "劉善總", "劉庭禎", "王荷若", "鄭竣文", 
    "李亞凡", "陳明志", "黃菁菁", "傅超俊", "陳昶安", "陳苡瑜", "許雅茹", "蔣宜蓁", 
    "劉筑昀", "胡瑋凡"
]

DOCTOR_ASSISTANT_MATCH = {
    "陳逸陽": "映璇", "王荷若": "萃屏", "李昆晏": "萃屏", "許雅茹": "和芸", "劉筑昀": "姿穎"
}
SATURDAY_SPECIAL_MATCH = {
    "官俊彥": "濘安",
    "陳明志": "菀庭"
}

ASSISTANTS = ["映璇", "和芸", "欣寧", "萃屏", "維珍", "菀庭", "姿穎", "濘安"]
COUNTER_PRIORITY = ["欣寧", "維珍", "和芸", "映璇", "姿穎"]

ONLY_COUNTER = ["欣寧"]
NO_COUNTER = ["萃屏", "菀庭", "濘安"]
NO_NIGHT_SHIFT = ["維珍"]

st.set_page_config(page_title="恩霖診所 - 自動排班系統 V11", layout="wide")

if "authenticated" not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.title("🔒 恩霖診所 - 排班系統登入")
    pwd = st.text_input("請輸入密碼：", type="password")
    if st.button("登入"):
        if pwd == "115": st.session_state.authenticated = True; st.rerun()
        else: st.error("密碼錯誤！")
    st.stop()
if "timeoff_db" not in st.session_state: st.session_state.timeoff_db = {}

st.title("🏥 恩霖診所 - 自動排班系統 V11")
st.info("🛡️ 已裝備「醫師白名單」完全忽略診間號碼，並升級「5/1 日期格式」精準掃描。")

tab1, tab2, tab3 = st.tabs(["📁 1. 上傳班表", "📝 2. 助理劃休", "🚀 3. AI 排班"])

with tab1:
    uploaded_file = st.file_uploader("上傳醫師班表 (Excel .xlsx)", type=["xlsx"])
    if uploaded_file: st.success("檔案讀取完成！")

with tab2:
    selected_ast = st.selectbox("選擇助理：", ASSISTANTS)
    if selected_ast not in st.session_state.timeoff_db:
        df = pd.DataFrame({"日期": [f"{i}號" for i in range(1, 32)], "早休": [False]*31, "午休": [False]*31, "晚休": [False]*31})
        if selected_ast in NO_NIGHT_SHIFT: df["晚休"] = True
        st.session_state.timeoff_db[selected_ast] = df
    edited_df = st.data_editor(st.session_state.timeoff_db[selected_ast], hide_index=True, key=f"ed_{selected_ast}")
    if st.button(f"儲存 {selected_ast} 休假"):
        st.session_state.timeoff_db[selected_ast] = edited_df
        st.toast(f"{selected_ast} 的休假已更新")

with tab3:
    if st.button("🚀 開始智慧排班", type="primary"):
        if not uploaded_file:
            st.error("請先上傳班表！")
        else:
            with st.spinner("AI 正在精準校對日期與醫師名單..."):
                try:
                    wb = openpyxl.load_workbook(uploaded_file)
                    ws = wb.active
                    
                    def is_on_leave(ast_name, d_num, shift_type):
                        if ast_name not in st.session_state.timeoff_db: return False
                        if d_num < 1 or d_num > 31: return False 
                        
                        db = st.session_state.timeoff_db[ast_name]
                        col_name = "早休" if "早" in shift_type else ("午休" if "午" in shift_type else "晚休")
                        try:
                            return bool(db[col_name].iloc[d_num - 1])
                        except Exception:
                            return False

                    shift_rows = {"早班": [], "午班": [], "晚班": []}
                    current_shift = None
                    for r in range(1, ws.max_row + 1):
                        val = str(ws.cell(row=r, column=1).value).replace(" ", "")
                        if "早班" in val: current_shift = "早班"
                        elif "午班" in val: current_shift = "午班"
                        elif "晚班" in val: current_shift = "晚班"
                        if current_shift: shift_rows[current_shift].append(r)

                    for col_idx in range(2, ws.max_column, 2):
                        day_num = -1
                        is_saturday = False
                        
                        # 📅 真實日期與星期掃描器：同時看 col_idx 與右邊一格，確保不漏掉 5/1 與 (五)
                        for r_check in range(1, 15): 
                            val1 = ws.cell(row=r_check, column=col_idx).value
                            val2 = ws.cell(row=r_check, column=col_idx+1).value
                            str1 = str(val1).replace(" ", "") if val1 else ""
                            str2 = str(val2).replace(" ", "") if val2 else ""
                            
                            if "六" in str1 or "六" in str2:
                                is_saturday = True
                            
                            for v in [val1, val2]:
                                if isinstance(v, datetime.datetime):
                                    day_num = v.day
                                elif isinstance(v, str):
                                    # 支援 5/1(五), 05/25, 2026-05-25 等各種格式
                                    m1 = re.search(r'\d{2,4}[-/]\d{1,2}[-/](\d{1,2})', v)
                                    m2 = re.search(r'\d{1,2}[-/](\d{1,2})', v)
                                    if m1: day_num = int(m1.group(1))
                                    elif m2: day_num = int(m2.group(1))

                        if day_num == -1:
                            day_num = col_idx // 2

                        for shift_name, rows in shift_rows.items():
                            working_now = set()
                            docs = []
                            
                            # A. 處理跟診與綁定
                            for r in rows:
                                d_name = ws.cell(row=r, column=col_idx).value
                                if d_name:
                                    clean_name = str(d_name).replace(" ", "")
                                    # 🛡️ 嚴格判斷：只處理「醫師白名單」內的人，直接無視診間號碼
                                    doc_found = next((doc for doc in VALID_DOCTORS if doc in clean_name), None)
                                    
                                    if doc_found and "休" not in clean_name:
                                        docs.append((r, doc_found))
                                        assigned_ast = None
                                        
                                        if is_saturday and doc_found in SATURDAY_SPECIAL_MATCH:
                                            assigned_ast = SATURDAY_SPECIAL_MATCH[doc_found]
                                        elif doc_found in DOCTOR_ASSISTANT_MATCH:
                                            assigned_ast = DOCTOR_ASSISTANT_MATCH[doc_found]
                                        
                                        if assigned_ast:
                                            ws.cell(row=r, column=col_idx+1).value = assigned_ast
                                            working_now.add(assigned_ast)
                            
                            if not docs: continue

                            # B. 櫃檯分配
                            counter_needed = 2 if len(docs) >= 4 else 1
                            assigned_counters = []
                            for candidate in COUNTER_PRIORITY:
                                if len(assigned_counters) >= counter_needed: break
                                if candidate not in working_now and not is_on_leave(candidate, day_num, shift_name):
                                    assigned_counters.append(candidate)
                                    working_now.add(candidate)
                            
                            while len(assigned_counters) < counter_needed:
                                pool = [a for a in ASSISTANTS if a not in working_now and a not in NO_COUNTER]
                                if pool:
                                    extra = pool[0]; assigned_counters.append(extra); working_now.add(extra)
                                else:
                                    assigned_counters.append("【缺人】"); break

                            c_idx = 0
                            for r in rows:
                                tag = str(ws.cell(row=r, column=1).value)
                                if "櫃" in tag and c_idx < len(assigned_counters):
                                    ws.cell(row=r, column=col_idx+1).value = assigned_counters[c_idx]
                                    c_idx += 1

                            # C. 其餘醫師跟診
                            for r, doc_found in docs:
                                if not ws.cell(row=r, column=col_idx+1).value: 
                                    pool = [a for a in ASSISTANTS if a not in working_now and a not in ONLY_COUNTER and a != "濘安"]
                                    if "晚" in shift_name and "維珍" in pool: pool.remove("維珍")
                                    
                                    if pool:
                                        chosen = pool[0]
                                        ws.cell(row=r, column=col_idx+1).value = chosen
                                        working_now.add(chosen)
                                    else:
                                        ws.cell(row=r, column=col_idx+1).value = "【缺人】"

                    output = io.BytesIO()
                    wb.save(output)
                    st.success("排班完成！已完美掃描正式班表格式。")
                    st.download_button("📥 下載 V11 終極穩定版", output.getvalue(), "恩霖診所_正式排班_V11.xlsx")
                except Exception as e:
                    st.error(f"錯誤：{e}")
