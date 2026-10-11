import io
import pandas as pd
import streamlit as st
from ortools.sat.python import cp_model

st.set_page_config(page_title="シフト自動作成システム", layout="wide")
st.title("📅 シフト自動作成システム")

# =====================================================================
# 🔒 パスワードロック機能（初期パスワード: 0927）
# =====================================================================
try:
    PASSWORD = st.secrets["APP_PASSWORD"]
except Exception:
    PASSWORD = "0927"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.subheader("ログイン")
    pwd = st.text_input("🔒 パスワードを入力してください", type="password")
    if st.button("ログイン", type="primary"):
        if pwd == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    st.stop()  # ログインするまで、これより下の画面やサイドバーを一切表示しない

# =====================================================================
# 基本設定・定数
# =====================================================================
S_HAYA, S_Q, S_OSO, S_IRI, S_AKE, S_YASU, S_YUKU, S_KEN, S_FREE = range(9)
shift_str = {
    S_HAYA: '早', S_Q: 'Q', S_OSO: '遅', S_IRI: '入', 
    S_AKE: '明', S_YASU: '休', S_YUKU: '有給', S_KEN: '研修', S_FREE: 'フリー'
}
str_to_shift = {v: k for k, v in shift_str.items()}

staffs = 16
staff_names = [chr(ord('A') + i) for i in range(staffs)]

# プルダウンで選べる選択肢
shift_options = ["", "早", "Q", "遅", "入", "明", "休", "有給", "研修", "フリー"]

# =====================================================================
# サイドバー（毎月の設定エリア）
# =====================================================================
st.sidebar.header("⚙️ 毎月の基本設定")
TARGET_DAYS = st.sidebar.number_input("月の日数", min_value=28, max_value=31, value=31)
TARGET_YASU = st.sidebar.number_input("全員が必ずとる「休」の日数", min_value=1, max_value=15, value=9)

default_q_days = [d for d in [1, 3, 6, 7, 8, 13, 14, 15, 17, 20, 22, 24, 27, 28, 29] if d <= TARGET_DAYS]
Q_SHIFT_DAYS = st.sidebar.multiselect(
    "「Q」シフトを【配置しない】日",
    options=list(range(1, TARGET_DAYS + 1)),
    default=default_q_days
)

mode = st.sidebar.radio(
    "実行モード",
    ["通常作成モード（自動3段階緩和）", "原因調査モード（デバッグ用・強制出力）"]
)

st.sidebar.divider()
if st.sidebar.button("🔓 ログアウト"):
    st.session_state.authenticated = False
    st.rerun()

# =====================================================================
# メイン画面：希望休・指定シフトの入力表（プルダウン式）
# =====================================================================
st.subheader("1. 希望休・指定シフトの入力")
st.caption("表のマス目をクリックすると、プルダウンからシフトを選択できます。空欄の場所はAIが自動で埋めます。")

init_df = pd.DataFrame("", index=staff_names, columns=[str(i) for i in range(1, TARGET_DAYS + 1)])

column_configuration = {
    str(i): st.column_config.SelectboxColumn(
        label=str(i),
        options=shift_options,
        default="",
        width="small"
    )
    for i in range(1, TARGET_DAYS + 1)
}

edited_df = st.data_editor(
    init_df,
    column_config=column_configuration,
    use_container_width=True,
    height=580
)

SPECIFIC_SHIFTS = {}
for s_idx, s_name in enumerate(staff_names):
    for d in range(1, TARGET_DAYS + 1):
        val = edited_df.loc[s_name, str(d)]
        if pd.notna(val) and str(val).strip() != "":
            val_str = str(val).strip()
            if val_str in str_to_shift:
                SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]

# =====================================================================
# 事前チェック関数
# =====================================================================
def check_shortage():
    total_slots = staffs * TARGET_DAYS
    work_slots = (3 + 3 + 2 + 2) * TARGET_DAYS + (TARGET_DAYS - len(Q_SHIFT_DAYS))
    specific_other_slots = sum(1 for (s, d), t in SPECIFIC_SHIFTS.items() if s < staffs and t not in [S_YASU, S_HAYA, S_OSO, S_IRI, S_AKE, S_Q])
    
    available_slots = total_slots - work_slots - specific_other_slots
    required_yasu = staffs * TARGET_YASU
    shortage = required_yasu - available_slots
    
    if shortage > 0:
        st.error(f"❌ 【重大なエラー：枠不足】指定された条件では、全員に「休」を{TARGET_YASU}回配置することが物理的に不可能です。（必要な休の枠に対して 【 {shortage} 枠 】 足りません）")
        return False
        
    for s in range(staffs):
        p_other = sum(1 for (s_idx, d_idx), t_idx in SPECIFIC_SHIFTS.items() if s_idx == s and t_idx not in [S_YASU, S_HAYA, S_OSO, S_IRI, S_AKE, S_Q])
        if TARGET_YASU - (TARGET_DAYS - p_other) > 0:
            st.error(f"❌ 【重大なエラー：個人の枠不足】{staff_names[s]}さんの有給・研修等の指定が多すぎるため、「休」を{TARGET_YASU}回配置できません。（【 {TARGET_YASU - (TARGET_DAYS - p_other)} 枠 】 足りません）")
            return False
    return True

# =====================================================================
# 計算ロジック（通常モード ＆ デバッグモード ※時間制限なし）
# =====================================================================
def run_normal_mode():
    q_days_idx = [d - 1 for d in Q_SHIFT_DAYS]
    
    def attempt_solve(consecutive_limit, strict_rank):
        model = cp_model.CpModel()
        B = {(s, d, t): model.NewBoolVar(f'B_{s}_{d}_{t}') for s in range(staffs) for d in range(TARGET_DAYS) for t in range(9)}
        assumptions = []
        assumption_dict = {}

        def add_enforced(constraint, name):
            b = model.NewBoolVar(name)
            constraint.OnlyEnforceIf(b)
            assumptions.append(b)
            assumption_dict[b.Index()] = name

        for s in range(staffs):
            for d in range(TARGET_DAYS):
                add_enforced(model.Add(sum(B[s, d, t] for t in range(9)) == 1), f"{staff_names[s]}さんの{d+1}日目は1シフトのみ")
                req_shift = SPECIFIC_SHIFTS.get((s, d+1))
                if req_shift != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
                if req_shift != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
                if req_shift is not None:
                    add_enforced(model.Add(B[s, d, req_shift] == 1), f"{staff_names[s]}さんの{d+1}日目は「{shift_str[req_shift]}」指定")
            
            add_enforced(model.Add(sum(B[s, d, S_YASU] for d in range(TARGET_DAYS)) == TARGET_YASU), f"{staff_names[s]}さんの「休」が{TARGET_YASU}日ではない")

        for d in range(TARGET_DAYS):
            add_enforced(model.Add(sum(B[s, d, S_HAYA] for s in range(staffs)) == 3), f"{d+1}日目の早番は3人")
            add_enforced(model.Add(sum(B[s, d, S_OSO]  for s in range(staffs)) == 3), f"{d+1}日目の遅番は3人")
            add_enforced(model.Add(sum(B[s, d, S_IRI]  for s in range(staffs)) == 2), f"{d+1}日目の入は2人")
            add_enforced(model.Add(sum(B[s, d, S_AKE]  for s in range(staffs)) == 2), f"{d+1}日目の明は2人")
            if d in q_days_idx:
                add_enforced(model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 0), f"{d+1}日目のQは0人")
            else:
                add_enforced(model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 1), f"{d+1}日目のQは1人")

        for s in range(staffs):
            for d in range(TARGET_DAYS - 1):
                add_enforced(model.AddImplication(B[s, d, S_IRI], B[s, d+1, S_AKE]), f"{staff_names[s]}さん: 入の次は明")
                add_enforced(model.AddImplication(B[s, d+1, S_AKE], B[s, d, S_IRI]), f"{staff_names[s]}さん: 明の前は入")
                add_enforced(model.AddImplication(B[s, d, S_AKE], B[s, d+1, S_YASU]), f"{staff_names[s]}さん: 明の次は休")
                if strict_rank:
                    add_enforced(model.AddImplication(B[s, d, S_Q], B[s, d+1, S_HAYA].Not()), f"{staff_names[s]}さん: Qの次早禁止")
                    add_enforced(model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_HAYA].Not()), f"{staff_names[s]}さん: 遅の次早禁止")
                    add_enforced(model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_Q].Not()), f"{staff_names[s]}さん: 遅の次Q禁止")

        for s in range(staffs):
            for d in range(TARGET_DAYS - consecutive_limit + 1):
                add_enforced(model.Add(sum(B[s, d+k, S_YASU] + B[s, d+k, S_YUKU] for k in range(consecutive_limit)) >= 1), f"{staff_names[s]}さん: {consecutive_limit}日間に1回は休または有給が必要")

        add_enforced(model.Add(sum(B[0, d, S_IRI] for d in range(TARGET_DAYS)) <= 2), "Aさんの「入」は2回まで")
        add_enforced(model.Add(sum(B[1, d, S_IRI] for d in range(TARGET_DAYS)) <= 2), "Bさんの「入」は2回まで")
        add_enforced(model.Add(sum(B[2, d, S_IRI] for d in range(TARGET_DAYS)) <= 3), "Cさんの「入」は3回まで")
        add_enforced(model.Add(sum(B[4, d, S_IRI] for d in range(TARGET_DAYS)) <= 3), "Eさんの「入」は3回まで")
        add_enforced(model.Add(sum(B[3, d, S_HAYA] for d in range(TARGET_DAYS)) == 0), "Dさんの「早」は禁止")
        add_enforced(model.Add(sum(B[9, d, S_IRI] for d in range(TARGET_DAYS)) == 0), "Jさんの「入」は禁止")

        iri_counts = [model.NewIntVar(0, TARGET_DAYS, f'iri_{s}') for s in range(staffs)]
        for s in range(staffs): model.Add(iri_counts[s] == sum(B[s, d, S_IRI] for d in range(TARGET_DAYS)))
        max_iri = model.NewIntVar(0, TARGET_DAYS, 'max_iri')
        model.AddMaxEquality(max_iri, iri_counts)
        model.Minimize(max_iri)

        model.AddAssumptions(assumptions)
        solver = cp_model.CpSolver()
        status = solver.Solve(model)
        return status, solver, B, assumption_dict, consecutive_limit, strict_rank

    with st.spinner("シフト表を計算中...（ステップ1: 横5マス休 / 順位制約あり）"):
        status, solver, B, assumption_dict, used_limit, used_strict = attempt_solve(5, True)
    if status == cp_model.INFEASIBLE:
        st.warning("⚠️ 横5マス制約では組めませんでした。横6マスに緩和して再計算します...")
        with st.spinner("再計算中...（ステップ2: 横6マス休 / 順位制約あり）"):
            status, solver, B, assumption_dict, used_limit, used_strict = attempt_solve(6, True)
    if status == cp_model.INFEASIBLE:
        st.warning("⚠️ 横6マスでも組めませんでした。順位制約を解除して再計算します...")
        with st.spinner("再計算中...（ステップ3: 横6マス休 / 順位制約緩和）"):
            status, solver, B, assumption_dict, used_limit, used_strict = attempt_solve(6, False)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        st.success(f"✅ シフトの作成が完了しました！（適用条件：横{used_limit}マス休 / 順位制約 {'あり' if used_strict else '緩和'}）")
        table_data = [[shift_str[next(t for t in range(9) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)]
        return pd.DataFrame(table_data, index=staff_names, columns=[str(i) for i in range(1, TARGET_DAYS + 1)])
    elif status == cp_model.INFEASIBLE:
        st.error("❌ 最大まで緩和しましたが、条件が厳しすぎて組めません。左のメニューから「原因調査モード」に切り替えて実行してみてください。")
        with st.expander("衝突している制約の詳細を見る"):
            for idx in solver.SufficientAssumptionsForInfeasibility():
                if idx in assumption_dict: st.write(f"・{assumption_dict[idx]}")
        return pd.DataFrame()
    else:
        st.error("❌ 計算を完了できませんでした。")
        return pd.DataFrame()

def run_debug_mode():
    model = cp_model.CpModel()
    q_days_idx = [d - 1 for d in Q_SHIFT_DAYS]
    B = {(s, d, t): model.NewBoolVar(f'B_{s}_{d}_{t}') for s in range(staffs) for d in range(TARGET_DAYS) for t in range(9)}
    penalties = []
    violations_info = []

    def add_soft(constraint, weight, desc):
        viol = model.NewBoolVar('')
        constraint.OnlyEnforceIf(viol.Not())
        penalties.append(viol * weight)
        violations_info.append((viol, desc))

    for s in range(staffs):
        for d in range(TARGET_DAYS):
            model.AddExactlyOne([B[s, d, t] for t in range(9)])
            req_shift = SPECIFIC_SHIFTS.get((s, d+1))
            if req_shift != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
            if req_shift != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
            if req_shift is not None: model.Add(B[s, d, req_shift] == 1)

    for s in range(staffs):
        add_soft(model.Add(sum(B[s, d, S_YASU] for d in range(TARGET_DAYS)) == TARGET_YASU), 50, f"【休数エラー】{staff_names[s]}さんの「休」が{TARGET_YASU}回ではありません")

    for d in range(TARGET_DAYS):
        add_soft(model.Add(sum(B[s, d, S_HAYA] for s in range(staffs)) == 3), 20, f"【人数エラー】{d+1}日目の「早番」が3人ではありません")
        add_soft(model.Add(sum(B[s, d, S_OSO]  for s in range(staffs)) == 3), 20, f"【人数エラー】{d+1}日目の「遅番」が3人ではありません")
        add_soft(model.Add(sum(B[s, d, S_IRI]  for s in range(staffs)) == 2), 20, f"【人数エラー】{d+1}日目の「入」が2人ではありません")
        add_soft(model.Add(sum(B[s, d, S_AKE]  for s in range(staffs)) == 2), 20, f"【人数エラー】{d+1}日目の「明」が2人ではありません")
        if d in q_days_idx:
            add_soft(model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 0), 20, f"【人数エラー】{d+1}日目に「Q」が配置されています")
        else:
            add_soft(model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 1), 20, f"【人数エラー】{d+1}日目の「Q」が1人ではありません")

    for s in range(staffs):
        for d in range(TARGET_DAYS - 1):
            add_soft(model.Add(B[s, d+1, S_AKE] >= B[s, d, S_IRI]), 50, f"【リレー違反】{staff_names[s]}さんの{d+1}日目「入」の次が「明」ではない")
            add_soft(model.Add(B[s, d, S_IRI] >= B[s, d+1, S_AKE]), 50, f"【リレー違反】{staff_names[s]}さんの{d+2}日目「明」の前が「入」ではない")
            add_soft(model.Add(B[s, d+1, S_YASU] >= B[s, d, S_AKE]), 50, f"【リレー違反】{staff_names[s]}さんの{d+1}日目「明」の次が「休」ではない")
            add_soft(model.Add(B[s, d, S_Q] + B[s, d+1, S_HAYA] <= 1), 10, f"【順位違反】{staff_names[s]}さんの{d+1}日目「Q」の次に「早」")
            add_soft(model.Add(B[s, d, S_OSO] + B[s, d+1, S_HAYA] <= 1), 10, f"【順位違反】{staff_names[s]}さんの{d+1}日目「遅」の次に「早」")
            add_soft(model.Add(B[s, d, S_OSO] + B[s, d+1, S_Q] <= 1), 10, f"【順位違反】{staff_names[s]}さんの{d+1}日目「遅」の次に「Q」")

    for s in range(staffs):
        for d in range(TARGET_DAYS - 4):
            add_soft(model.Add(sum(B[s, d+k, S_YASU] + B[s, d+k, S_YUKU] for k in range(5)) >= 1), 5, f"【連勤超過】{staff_names[s]}さんが{d+1}日目から連続勤務オーバー(横5マス以内に休/有給なし)")

    add_soft(model.Add(sum(B[0, d, S_IRI] for d in range(TARGET_DAYS)) <= 2), 20, "【回数超過】Aさんの「入」は2回を超えました")
    add_soft(model.Add(sum(B[1, d, S_IRI] for d in range(TARGET_DAYS)) <= 2), 20, "【回数超過】Bさんの「入」は2回を超えました")
    add_soft(model.Add(sum(B[2, d, S_IRI] for d in range(TARGET_DAYS)) <= 3), 20, "【回数超過】Cさんの「入」は3回を超えました")
    add_soft(model.Add(sum(B[4, d, S_IRI] for d in range(TARGET_DAYS)) <= 3), 20, "【回数超過】Eさんの「入」は3回を超えました")
    add_soft(model.Add(sum(B[3, d, S_HAYA] for d in range(TARGET_DAYS)) == 0), 20, "【禁止配置】Dさんに「早」が配置されています")
    add_soft(model.Add(sum(B[9, d, S_IRI] for d in range(TARGET_DAYS)) == 0), 20, "【禁止配置】Jさんに「入」が配置されています")

    model.Minimize(sum(penalties))
    solver = cp_model.CpSolver()
    
    with st.spinner("デバッグモードで強制出力しています..."):
        status = solver.Solve(model)
        
    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        broken_rules = [desc for viol_var, desc in violations_info if solver.Value(viol_var)]
        if not broken_rules:
            st.success("✅ ルール違反は一つもありませんでした！")
        else:
            st.error(f"🚨 以下のルールを破って無理やり作成しました（計 {len(broken_rules)} 件の違反）：")
            for rule in broken_rules:
                st.write(f"・ {rule}")
        table_data = [[shift_str[next(t for t in range(9) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)]
        return pd.DataFrame(table_data, index=staff_names, columns=[str(i) for i in range(1, TARGET_DAYS + 1)])
    else:
        st.error("希望休などの絶対ルール同士が矛盾しています。")
        return pd.DataFrame()

# =====================================================================
# 実行ボタンと結果出力
# =====================================================================
st.subheader("2. シフトの作成と出力")
if st.button("🚀 シフト表を作成する", type="primary", use_container_width=True):
    if check_shortage():
        if "通常作成" in mode:
            result_df = run_normal_mode()
        else:
            result_df = run_debug_mode()
            
        if not result_df.empty:
            st.dataframe(result_df, use_container_width=True, height=580)
            
            # Excelダウンロード機能
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name='シフト表')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 完成したシフト表をExcelでダウンロード",
                data=excel_data,
                file_name="shift_schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )