import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font
import datetime

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

# 醫師固定跟診助理（平日）
DOCTOR_ASSISTANT_MATCH = {
    "陳逸陽": "映璇",
    "王荷若": "萃屏",
    "李昆晏": "萃屏",
    "許雅茹": "和芸",
    "劉筑昀": "姿穎",
}

# 星期六特殊綁定
SATURDAY_SPECIAL_MATCH = {
    "官俊彥": "濘安",
    "陳明志": "菀庭",
}

# 助理名單（詩庭產假、怡婷為主管，皆不排班）
ASSISTANTS = ["映璇", "和芸", "欣寧", "萃屏", "維珍", "菀庭", "姿穎", "濘安"]

# 櫃檯優先順序
COUNTER_PRIORITY = ["欣寧", "維珍", "和芸", "映璇", "姿穎"]

ONLY_COUNTER = ["欣寧"]        # 只做櫃檯
NO_COUNTER   = ["萃屏", "菀庭", "濘安"]  # 不做櫃檯
NO_NIGHT_SHIFT = ["維珍"]      # 不排晚班

# ==========================================
# 2. Streamlit 設定
# ==========================================
st.set_page_config(page_title="恩霖診所 - 自動排班系統", layout="wide")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 恩霖診所 - 排班系統登入")
    pwd = st.text_input("請輸入密碼：", type="password")
    if st.button("登入"):
        if pwd == "115":
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("密碼錯誤！")
    st.stop()

if "timeoff_db" not in st.session_state:
    st.session_state.timeoff_db = {}

st.title("🏥 恩霖診所 - 自動排班系統（修正版）")

tab1, tab2, tab3 = st.tabs(["📁 1. 上傳班表", "📝 2. 助理劃休", "🚀 3. AI 排班"])

# ==========================================
# Tab 1: 上傳班表
# ==========================================
with tab1:
    uploaded_file = st.file_uploader("上傳醫師班表 (Excel .xlsx)", type=["xlsx"])
    if uploaded_file:
        st.success("檔案讀取完成！")

# ==========================================
# Tab 2: 助理劃休
# ==========================================

# 星期六對應的日期（從上傳班表讀取；若未上傳則預設空集合）
# 直接用固定規則：5月2,9,16,23,30 為六
SATURDAY_DATES = {2, 9, 16, 23, 30}

with tab2:
    selected_ast = st.selectbox("選擇助理：", ASSISTANTS)

    # 初始化資料（若該助理尚無記錄）
    if selected_ast not in st.session_state.timeoff_db:
        df = pd.DataFrame({
            "日期":   [f"{i}號{'（六）' if i in SATURDAY_DATES else ''}" for i in range(1, 32)],
            "休整天": [False] * 31,
            "早休":   [False] * 31,
            "午休":   [False] * 31,
            "晚休":   [False] * 31,
        })
        if selected_ast in NO_NIGHT_SHIFT:
            df["晚休"] = True
        st.session_state.timeoff_db[selected_ast] = df

    # 已儲存的資料
    saved_df = st.session_state.timeoff_db[selected_ast].copy()

    # 欄位設定
    column_config = {
        "日期": st.column_config.TextColumn("日期", width="small", disabled=True),
        "休整天": st.column_config.CheckboxColumn("✅ 休整天", width="small",
            help="打勾後自動勾選早、午、晚休"),
        "早休": st.column_config.CheckboxColumn("早休", width="small"),
        "午休": st.column_config.CheckboxColumn("午休", width="small"),
        "晚休": st.column_config.CheckboxColumn("晚休", width="small"),
    }

    # 禁用星期六晚休、NO_NIGHT_SHIFT 助理晚休
    disabled_cols = []
    if selected_ast in NO_NIGHT_SHIFT:
        disabled_cols.append("晚休")

    st.info("📌 星期六（2、9、16、23、30 號）無晚班，即使勾選晚休也不會計入排班。")

    edited_df = st.data_editor(
        saved_df,
        hide_index=True,
        key=f"ed_{selected_ast}",
        column_config=column_config,
        disabled=disabled_cols,
        use_container_width=True,
    )

    # 「休整天」打勾 → 自動帶動早午晚全勾
    for i, row in edited_df.iterrows():
        day_num = i + 1
        if row["休整天"]:
            edited_df.at[i, "早休"] = True
            edited_df.at[i, "午休"] = True
            if day_num not in SATURDAY_DATES and selected_ast not in NO_NIGHT_SHIFT:
                edited_df.at[i, "晚休"] = True
        # 星期六晚休強制 False
        if day_num in SATURDAY_DATES:
            edited_df.at[i, "晚休"] = False

    # 儲存按鈕
    col_btn, col_status = st.columns([1, 3])
    with col_btn:
        save_clicked = st.button(f"💾 儲存 {selected_ast} 休假", type="primary")

    if save_clicked:
        st.session_state.timeoff_db[selected_ast] = edited_df.copy()
        with col_status:
            # 統計已勾選的休假
            total_early = edited_df["早休"].sum()
            total_noon  = edited_df["午休"].sum()
            total_night = edited_df["晚休"].sum()
            full_days   = edited_df["休整天"].sum()
            st.success(
                f"✅ {selected_ast} 休假已儲存！　"
                f"休整天：{full_days} 天　早休：{total_early} 次　"
                f"午休：{total_noon} 次　晚休：{total_night} 次"
            )

    # 視覺提示：顯示目前已儲存的休假摘要（與 editor 分開顯示，避免混淆）
    st.divider()
    st.caption("📌 目前已儲存的休假紀錄")
    saved = st.session_state.timeoff_db[selected_ast]
    full_day_list  = [f"{i+1}號" for i, r in saved.iterrows() if r["休整天"]]
    early_list     = [f"{i+1}號" for i, r in saved.iterrows() if r["早休"] and not r["休整天"]]
    noon_list      = [f"{i+1}號" for i, r in saved.iterrows() if r["午休"] and not r["休整天"]]
    night_list     = [f"{i+1}號" for i, r in saved.iterrows() if r["晚休"] and not r["休整天"]]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown("**🔴 休整天**")
        st.write("、".join(full_day_list) if full_day_list else "（無）")
    with c2:
        st.markdown("**🟠 僅早休**")
        st.write("、".join(early_list) if early_list else "（無）")
    with c3:
        st.markdown("**🟡 僅午休**")
        st.write("、".join(noon_list) if noon_list else "（無）")
    with c4:
        st.markdown("**🟤 僅晚休**")
        st.write("、".join(night_list) if night_list else "（無）")

# ==========================================
# Tab 3: 排班
# ==========================================
with tab3:
    if st.button("🚀 開始智慧排班", type="primary"):
        if not uploaded_file:
            st.error("請先上傳班表！")
        else:
            with st.spinner("正在解析班表並排班中..."):
                try:
                    wb = openpyxl.load_workbook(uploaded_file)
                    ws = wb.active

                    # ── 輔助函式 ──────────────────────────────────────

                    def is_on_leave(ast_name, day_num, shift_type):
                        """查詢助理當天該班是否請假"""
                        if ast_name not in st.session_state.timeoff_db:
                            return False
                        if not (1 <= day_num <= 31):
                            return False
                        db = st.session_state.timeoff_db[ast_name]
                        col = "早休" if "早" in shift_type else ("午休" if "午" in shift_type else "晚休")
                        try:
                            return bool(db[col].iloc[day_num - 1])
                        except Exception:
                            return False

                    def write_cell(row, col, value):
                        """寫入儲存格並套用統一字型"""
                        cell = ws.cell(row, col)
                        cell.value = value
                        cell.font = CELL_FONT

                    def find_doctors_in_cell(cell_value):
                        """從儲存格內容找出所有醫師名（支援「官+王」合診格式）"""
                        if not cell_value:
                            return []
                        text = str(cell_value).replace(" ", "")
                        if "休" in text:
                            return []
                        return [doc for doc in VALID_DOCTORS if doc in text]

                    def col_to_day_index(col, date_cols):
                        """
                        根據欄號找出對應的 (day_num, is_saturday, asst_col)。
                        date_cols = [(doc_col, asst_col, day_num, is_saturday), ...]
                        """
                        for dc, ac, dn, is_sat in date_cols:
                            if col == dc:
                                return ac, dn, is_sat
                        return None, None, None

                    # ── 找出所有週塊起始行 ────────────────────────────
                    enlin_rows = []
                    for r in range(1, ws.max_row + 1):
                        v = ws.cell(r, 1).value
                        if v and "恩霖" in str(v):
                            enlin_rows.append(r)

                    if not enlin_rows:
                        st.error("找不到班表週塊（含「恩霖」的列），請確認上傳的檔案格式正確。")
                        st.stop()

                    # ── 處理每個週塊 ──────────────────────────────────
                    for week_idx, week_start in enumerate(enlin_rows):
                        week_end = enlin_rows[week_idx + 1] if week_idx + 1 < len(enlin_rows) else ws.max_row + 1

                        # 1. 讀取日期列：找出 (doc_col, asst_col, day_num, is_saturday)
                        #    規則：偶數欄放醫師，奇數欄放助理，日期/星期交替排在日期列
                        date_cols = []
                        date_row = week_start
                        for c in range(2, ws.max_column + 1, 2):
                            v = ws.cell(date_row, c).value
                            v_next = ws.cell(date_row, c + 1).value  # 星期
                            if isinstance(v, datetime.datetime):
                                day_num = v.day
                                is_sat = str(v_next).strip() == "六" if v_next else False
                                date_cols.append((c, c + 1, day_num, is_sat))

                        if not date_cols:
                            continue  # 此週塊無日期，跳過

                        # 2. 找出班別區段的列號
                        #    早班：早班_label+1 ~ 早櫃2-1
                        #    午班：早櫃1+2 ~ 午櫃2-1  (跳過緊接在早櫃1後的特殊醫師label列)
                        #    晚班：晚班_label+1 ~ 晚櫃2-1
                        row_labels = {}
                        for r in range(week_start, week_end):
                            v1 = str(ws.cell(r, 1).value or "").replace(" ", "")
                            if "早班" in v1:   row_labels.setdefault("早班",   r)
                            elif "早櫃2" in v1: row_labels.setdefault("早櫃2", r)
                            elif "早櫃1" in v1: row_labels.setdefault("早櫃1", r)
                            elif "午班" in v1:  row_labels.setdefault("午班",  r)
                            elif "午櫃2" in v1: row_labels.setdefault("午櫃2", r)
                            elif "午櫃1" in v1: row_labels.setdefault("午櫃1", r)
                            elif "晚班" in v1:  row_labels.setdefault("晚班",  r)
                            elif "晚櫃2" in v1: row_labels.setdefault("晚櫃2", r)
                            elif "晚櫃1" in v1: row_labels.setdefault("晚櫃1", r)
                            elif "備註" in v1:  row_labels.setdefault("備註",  r); break

                        # 定義三個班別的醫師行範圍與對應的櫃檯行
                        shifts = []

                        if "早班" in row_labels and "早櫃2" in row_labels:
                            doc_rows = [r for r in range(row_labels["早班"] + 1, row_labels["早櫃2"])
                                        if str(ws.cell(r, 1).value or "").strip() in ("11", "21", "22", "23", "24")]
                            counter_rows = [row_labels.get("早櫃2"), row_labels.get("早櫃1")]
                            shifts.append(("早班", doc_rows, [r for r in counter_rows if r]))

                        if "早櫃1" in row_labels and "午櫃2" in row_labels:
                            # 午班醫師列：早櫃1 之後，到午櫃2 之前（跳過非數字列）
                            doc_rows = [r for r in range(row_labels["早櫃1"] + 1, row_labels["午櫃2"])
                                        if str(ws.cell(r, 1).value or "").strip() in ("11", "21", "22", "23", "24")]
                            counter_rows = [row_labels.get("午櫃2"), row_labels.get("午櫃1")]
                            shifts.append(("午班", doc_rows, [r for r in counter_rows if r]))

                        if "晚班" in row_labels and "晚櫃2" in row_labels:
                            doc_rows = [r for r in range(row_labels["晚班"] + 1, row_labels["晚櫃2"])
                                        if str(ws.cell(r, 1).value or "").strip() in ("11", "21", "22", "23", "24")]
                            counter_rows = [row_labels.get("晚櫃2"), row_labels.get("晚櫃1")]
                            shifts.append(("晚班", doc_rows, [r for r in counter_rows if r]))

                        # 3. 對每個班別 × 每一天 做排班
                        for shift_name, doc_rows, counter_rows in shifts:
                            for doc_col, asst_col, day_num, is_saturday in date_cols:

                                # 收集本班、本天的醫師清單
                                docs_this_shift = []
                                for r in doc_rows:
                                    doctors = find_doctors_in_cell(ws.cell(r, doc_col).value)
                                    for doc in doctors:
                                        docs_this_shift.append((r, doc))

                                if not docs_this_shift:
                                    continue  # 這天這班沒有醫師，跳過

                                working_now = set()

                                # A. 固定跟診綁定（已在班表填好的助理，直接納入 working_now）
                                for r, doc in docs_this_shift:
                                    existing_asst = ws.cell(r, asst_col).value
                                    if existing_asst and str(existing_asst).strip() in ASSISTANTS:
                                        working_now.add(str(existing_asst).strip())

                                # B. 對尚未填助理的醫師行，依規則填入
                                for r, doc in docs_this_shift:
                                    existing_asst = ws.cell(r, asst_col).value
                                    already_filled = existing_asst and str(existing_asst).strip() in ASSISTANTS

                                    if already_filled:
                                        continue  # 班表已寫好，不覆蓋

                                    assigned = None

                                    # 星期六特殊綁定
                                    if is_saturday and doc in SATURDAY_SPECIAL_MATCH:
                                        candidate = SATURDAY_SPECIAL_MATCH[doc]
                                        if candidate not in working_now and not is_on_leave(candidate, day_num, shift_name):
                                            assigned = candidate

                                    # 平日固定綁定
                                    if not assigned and doc in DOCTOR_ASSISTANT_MATCH:
                                        candidate = DOCTOR_ASSISTANT_MATCH[doc]
                                        if candidate not in working_now and not is_on_leave(candidate, day_num, shift_name):
                                            assigned = candidate

                                    # 一般助理池（排除晚班 維珍、排除 ONLY_COUNTER）
                                    if not assigned:
                                        pool = [
                                            a for a in ASSISTANTS
                                            if a not in working_now
                                            and a not in ONLY_COUNTER
                                            and a not in NO_COUNTER
                                            and not is_on_leave(a, day_num, shift_name)
                                        ]
                                        if "晚" in shift_name:
                                            pool = [a for a in pool if a not in NO_NIGHT_SHIFT]
                                        if pool:
                                            assigned = pool[0]

                                    if assigned:
                                        write_cell(r, asst_col, assigned)
                                        working_now.add(assigned)
                                    else:
                                        write_cell(r, asst_col, "缺")

                                # C. 櫃檯分配
                                counter_needed = 2 if len(docs_this_shift) >= 4 else 1
                                assigned_counters = []

                                for candidate in COUNTER_PRIORITY:
                                    if len(assigned_counters) >= counter_needed:
                                        break
                                    if (candidate not in working_now
                                            and not is_on_leave(candidate, day_num, shift_name)):
                                        assigned_counters.append(candidate)
                                        working_now.add(candidate)

                                # 補足缺口
                                while len(assigned_counters) < counter_needed:
                                    pool = [
                                        a for a in ASSISTANTS
                                        if a not in working_now and a not in NO_COUNTER
                                        and not is_on_leave(a, day_num, shift_name)
                                    ]
                                    if pool:
                                        assigned_counters.append(pool[0])
                                        working_now.add(pool[0])
                                    else:
                                        assigned_counters.append("缺")
                                        break

                                # 寫入櫃檯列
                                c_idx = 0
                                for cr in counter_rows:
                                    if c_idx < len(assigned_counters) and cr:
                                        write_cell(cr, asst_col, assigned_counters[c_idx])
                                        c_idx += 1

                    # ── 輸出 ──────────────────────────────────────────
                    output = io.BytesIO()
                    wb.save(output)
                    st.success("✅ 排班完成！")
                    st.download_button(
                        "📥 下載排班結果",
                        output.getvalue(),
                        "恩霖診所_排班結果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                except Exception as e:
                    st.error(f"排班過程發生錯誤：{e}")
                    import traceback
                    st.code(traceback.format_exc())
