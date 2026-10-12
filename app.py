import io
import calendar
import datetime
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
    st.stop()

# =====================================================================
# 基本設定・定数
# =====================================================================
S_HAYA, S_Q, S_OSO, S_IRI, S_AKE, S_YASU, S_YUKU, S_KEN, S_FREE = range(9)
shift_str = {
    S_HAYA: '早', S_Q: 'Q', S_OSO: '遅', S_IRI: '入', 
    S_AKE: '明', S_YASU: '休', S_YUKU: '有給', S_KEN: '研修', S_FREE: 'フリー'
}
# 「希望休」も内部計算では S_YASU（休）として扱う
str_to_shift = {v: k for k, v in shift_str.items()}
str_to_shift['希望休'] = S_YASU

shift_options = ["", "早", "Q", "遅", "入", "明", "希望休", "休", "有給", "研修", "フリー"]
limit_options = ["制限なし"] + [str(i) for i in range(0, 32)]
prev_work_options = [str(i) for i in range(0, 6)]  # 前月連勤の選択肢（0〜5日）
PREV_COL = "前月連勤"
WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]

# ★初期表示のスタッフ名リスト（画面上でいつでも個人名に書き換え・Excel保存が可能です）
DEFAULT_STAFF_NAMES = [chr(ord('A') + i) for i in range(16)]

def get_default_staff_df():
    data = []
    for i, name in enumerate(DEFAULT_STAFF_NAMES):
        haya_lim = "0" if i == 3 else "制限なし"  # 4人目(D)は早禁止
        q_lim = "制限なし"
        oso_lim = "制限なし"
        if i in [0, 1]: iri_lim = "2"             # 1, 2人目(A, B)は入2回まで
        elif i in [2, 4]: iri_lim = "3"           # 3, 5人目(C, E)は入3回まで
        elif i == 9: iri_lim = "0"                # 10人目(J)は入禁止
        else: iri_lim = "制限なし"
        data.append({"スタッフ名": name, "早の上限": haya_lim, "Qの上限": q_lim, "遅の上限": oso_lim, "入の上限": iri_lim})
    return pd.DataFrame(data)

def extract_day_num(col_name):
    s = str(col_name).strip()
    if "(" in s:
        s = s.split("(")[0].strip()
    return int(s) if s.isdigit() else None

if "staff_df" not in st.session_state:
    st.session_state.staff_df = get_default_staff_df()
if "req_df" not in st.session_state:
    init_cols = [PREV_COL] + [str(i) for i in range(1, 32)]
    init_req = pd.DataFrame("", index=st.session_state.staff_df["スタッフ名"].tolist(), columns=init_cols)
    init_req[PREV_COL] = "0"
    st.session_state.req_df = init_req
if "editor_key" not in st.session_state:
    st.session_state.editor_key = 0

# =====================================================================
# サイドバー（対象年月・基本設定・データの保存と復元）
# =====================================================================
st.sidebar.header("⚙️ 毎月の基本設定")

now = datetime.datetime.now()
col_y, col_m = st.sidebar.columns(2)
target_year = col_y.number_input("対象年", min_value=2024, max_value=2035, value=now.year)
target_month = col_m.number_input("対象月", min_value=1, max_value=12, value=now.month)

_, auto_days = calendar.monthrange(int(target_year), int(target_month))
TARGET_DAYS = st.sidebar.number_input("月の日数（自動計算）", min_value=28, max_value=31, value=auto_days)

day_labels = {
    d: f"{d}({WEEKDAYS_JP[datetime.date(int(target_year), int(target_month), d).weekday()]})"
    for d in range(1, TARGET_DAYS + 1)
}

TARGET_YASU = st.sidebar.number_input("全員が必ずとる「休＋希望休」の日数", min_value=1, max_value=15, value=9)

default_q_days = [d for d in [1, 3, 6, 7, 8, 13, 14, 15, 17, 20, 22, 24, 27, 28, 29] if d <= TARGET_DAYS]
Q_SHIFT_DAYS = st.sidebar.multiselect(
    "「Q」シフトを【配置しない】日",
    options=list(range(1, TARGET_DAYS + 1)),
    default=default_q_days,
    format_func=lambda d: day_labels.get(d, str(d))
)

mode = st.sidebar.radio(
    "実行モード",
    ["通常作成モード（高速・自動3段階緩和）", "原因調査モード（デバッグ用・強制出力）"]
)

st.sidebar.divider()
st.sidebar.subheader("📂 保存データの読み込み")
st.sidebar.caption("一時保存したExcelや、先月の完成シフト表Excelを読み込むと、個人名や条件がそのまま復元されます。")

uploaded_file = st.sidebar.file_uploader("Excelファイルを選択", type=["xlsx"])
if uploaded_file is not None:
    if st.sidebar.button("🔄 データを画面に反映する", type="primary", use_container_width=True):
        try:
            xls = pd.ExcelFile(uploaded_file)
            if "スタッフ設定" in xls.sheet_names:
                loaded_staff = pd.read_excel(xls, "スタッフ設定").fillna("制限なし").astype(str)
                st.session_state.staff_df = loaded_staff
            if "希望休入力" in xls.sheet_names:
                loaded_req = pd.read_excel(xls, "希望休入力", index_col=0).fillna("").astype(str)
                norm_cols = [PREV_COL] + [str(i) for i in range(1, 32)]
                normalized_req = pd.DataFrame("", index=loaded_req.index.astype(str), columns=norm_cols)
                normalized_req[PREV_COL] = "0"
                for col in loaded_req.columns:
                    if str(col).strip() == PREV_COL:
                        cleaned_prev = [str(int(float(v))) if str(v).replace('.', '', 1).isdigit() else "0" for v in loaded_req[col].values]
                        normalized_req[PREV_COL] = cleaned_prev
                    else:
                        d_num = extract_day_num(col)
                        if d_num and 1 <= d_num <= 31:
                            normalized_req[str(d_num)] = loaded_req[col].values
                st.session_state.req_df = normalized_req
            st.session_state.editor_key += 1
            st.sidebar.success("✅ データを復元しました！")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"読み込みエラー: ファイル形式を確認してください ({e})")

# =====================================================================
# 1. スタッフ名・個人条件の設定（個人名への書き換え対応）
# =====================================================================
st.subheader("1. スタッフ名・個人条件の設定")
with st.expander("👥 スタッフ名（個人名への書き換え）・シフト上限・禁止条件を編集する", expanded=True):
    st.caption("・左端の **「スタッフ名」** のマスをダブルクリックすると、アルファベットから個人名（例：佐藤、鈴木など）に自由に書き換えられます。\n・「0」にするとそのシフトは【禁止】になります（例：早の上限＝0 → 早番禁止）。")
    
    # Excel等から一括で名前を貼り付けられる便利機能
    with st.popover("📝 スタッフ名をまとめて一括変更する（コピペ入力）"):
        st.caption("Excelやメモ帳からスタッフ名を改行区切り（またはスペース・読点区切り）で貼り付けて「反映」を押すと、上から順番に個人名へ一括変更されます。")
        bulk_names_text = st.text_area("スタッフ名リストを入力（例：佐藤、鈴木、高橋...）", height=150)
        if st.button("名前を一括反映する", type="primary"):
            raw_list = [n.strip() for n in bulk_names_text.replace("、", "\n").replace(",", "\n").replace(" ", "\n").replace(" ", "\n").splitlines() if n.strip()]
            if raw_list:
                new_staff_df = st.session_state.staff_df.copy()
                old_req_df = st.session_state.req_df.copy()
                for idx, new_name in enumerate(raw_list):
                    if idx < len(new_staff_df):
                        new_staff_df.loc[idx, "スタッフ名"] = new_name
                    else:
                        new_row = pd.DataFrame([{"スタッフ名": new_name, "早の上限": "制限なし", "Qの上限": "制限なし", "遅の上限": "制限なし", "入の上限": "制限なし"}])
                        new_staff_df = pd.concat([new_staff_df, new_row], ignore_index=True)
                st.session_state.staff_df = new_staff_df
                # 希望休テーブルの行名も新しい名前に連動して引き継ぐ
                updated_names = new_staff_df["スタッフ名"].tolist()
                new_req_df = pd.DataFrame("", index=updated_names, columns=old_req_df.columns)
                new_req_df[PREV_COL] = "0"
                for i in range(min(len(old_req_df), len(updated_names))):
                    new_req_df.iloc[i] = old_req_df.iloc[i].values
                st.session_state.req_df = new_req_df
                st.session_state.editor_key += 1
                st.rerun()

    staff_col_config = {
        "スタッフ名": st.column_config.TextColumn("スタッフ名（クリックで編集）", required=True, width="medium"),
        "早の上限": st.column_config.SelectboxColumn("早の上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "Qの上限": st.column_config.SelectboxColumn("Qの上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "遅の上限": st.column_config.SelectboxColumn("遅の上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "入の上限": st.column_config.SelectboxColumn("入の上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
    }

    edited_staff_df = st.data_editor(
        st.session_state.staff_df,
        column_config=staff_col_config,
        num_rows="dynamic",
        use_container_width=True,
        height=min(420, 65 + len(st.session_state.staff_df) * 35),
        key=f"staff_editor_{st.session_state.editor_key}"
    )

raw_names = [str(n).strip() if pd.notna(n) and str(n).strip() != "" else f"スタッフ{i+1}" for i, n in enumerate(edited_staff_df["スタッフ名"])]
staff_names = []
for i, name in enumerate(raw_names):
    if name in staff_names:
        staff_names.append(f"{name}_{i+1}")
    else:
        staff_names.append(name)
staffs = len(staff_names)

# =====================================================================
# 2. メイン画面：希望休・指定シフトの入力表（前月連勤＋曜日付き）
# =====================================================================
st.subheader(f"2. 希望休・指定シフトの入力（{target_year}年{target_month}月）")
st.caption("・上の表で書き換えた個人名がそのまま左端に表示されます。\n・左端の **「前月連勤」** には、前月末時点での連続勤務日数（0〜5日）を選択してください。\n・「希望休」と「休」はどちらも同じ休みとして計算されますが、完成したシフト表には「希望休」と区別して表記されます。")

day_cols = [str(i) for i in range(1, TARGET_DAYS + 1)]
all_req_cols = [PREV_COL] + day_cols
current_req_df = pd.DataFrame("", index=staff_names, columns=all_req_cols)
current_req_df[PREV_COL] = "0"

prev_req = st.session_state.req_df
for i, s_name in enumerate(staff_names):
    if s_name in prev_req.index:
        for c in all_req_cols:
            if c in prev_req.columns:
                val = str(prev_req.loc[s_name, c]).strip()
                current_req_df.loc[s_name, c] = val if (c != PREV_COL or val in prev_work_options) else "0"
    elif i < len(prev_req.index):
        old_name = prev_req.index[i]
        for c in all_req_cols:
            if c in prev_req.columns:
                val = str(prev_req.loc[old_name, c]).strip()
                current_req_df.loc[s_name, c] = val if (c != PREV_COL or val in prev_work_options) else "0"

column_configuration = {
    PREV_COL: st.column_config.SelectboxColumn(
        label="前月連勤",
        help="前月末時点での連続勤務日数（0〜5）",
        options=prev_work_options,
        default="0",
        width="small",
        required=True
    )
}
for i in range(1, TARGET_DAYS + 1):
    column_configuration[str(i)] = st.column_config.SelectboxColumn(
        label=day_labels[i],
        options=shift_options,
        default="",
        width="small"
    )

edited_req_df = st.data_editor(
    current_req_df,
    column_config=column_configuration,
    use_container_width=True,
    height=min(620, 65 + staffs * 35),
    key=f"req_editor_{st.session_state.editor_key}"
)

st.session_state.staff_df = edited_staff_df
st.session_state.req_df = edited_req_df

save_req_df = edited_req_df.copy()
save_req_df.columns = [PREV_COL] + [day_labels[int(c)] for c in day_cols]

save_output = io.BytesIO()
with pd.ExcelWriter(save_output, engine='openpyxl') as writer:
    edited_staff_df.to_excel(writer, sheet_name='スタッフ設定', index=False)
    save_req_df.to_excel(writer, sheet_name='希望休入力')
save_excel_bytes = save_output.getvalue()

col_save1, col_save2 = st.columns([1, 1])
with col_save1:
    st.download_button(
        label="💾 現在の入力進捗（個人名・設定・前月連勤・希望休）をExcelに一時保存",
        data=save_excel_bytes,
        file_name=f"shift_progress_{target_year}_{target_month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.sidebar.download_button(
    label="💾 現在の入力進捗を保存",
    data=save_excel_bytes,
    file_name=f"shift_progress_{target_year}_{target_month:02d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.sidebar.divider()
if st.sidebar.button("🔓 ログアウト"):
    st.session_state.authenticated = False
    st.rerun()

# =====================================================================
# 画面の表から 前月連勤・指定シフト・希望休の場所 を読み取る
# =====================================================================
SPECIFIC_SHIFTS = {}
KIBOU_YASU_CELLS = set()
PREV_CONSECUTIVE = {}

for s_idx, s_name in enumerate(staff_names):
    p_val = str(edited_req_df.loc[s_name, PREV_COL]).strip()
    PREV_CONSECUTIVE[s_idx] = int(p_val) if p_val.isdigit() else 0
    
    for d in range(1, TARGET_DAYS + 1):
        val = edited_req_df.loc[s_name, str(d)]
        if pd.notna(val) and str(val).strip() != "":
            val_str = str(val).strip()
            if val_str in str_to_shift:
                SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]
                if val_str == "希望休":
                    KIBOU_YASU_CELLS.add((s_idx, d - 1))

def get_personal_limits():
    limits = []
    col_map = [("早の上限", S_HAYA, "早"), ("Qの上限", S_Q, "Q"), ("遅の上限", S_OSO, "遅"), ("入の上限", S_IRI, "入")]
    for s_idx in range(staffs):
        row = edited_staff_df.iloc[s_idx]
        s_name = staff_names[s_idx]
        for col_name, s_type, s_label in col_map:
            val = str(row[col_name]).strip() if pd.notna(row[col_name]) else "制限なし"
            if val != "制限なし" and val.isdigit():
                max_c = int(val)
                desc = f"{s_name}さんの「{s_label}」は禁止(0回)" if max_c == 0 else f"{s_name}さんの「{s_label}」は{max_c}回まで"
                limits.append((s_idx, s_type, max_c, desc))
    return limits

# =====================================================================
# 完成したシフト表の整形（希望休の復元・右端への回数集計列の追加・色分け）
# =====================================================================
def format_result_df(raw_table_data):
    for s_idx, d_idx in KIBOU_YASU_CELLS:
        if 0 <= s_idx < len(raw_table_data) and 0 <= d_idx < TARGET_DAYS:
            if raw_table_data[s_idx][d_idx] == '休':
                raw_table_data[s_idx][d_idx] = '希望休'
    cols = [day_labels[i] for i in range(1, TARGET_DAYS + 1)]
    df = pd.DataFrame(raw_table_data, index=staff_names, columns=cols)
    
    df["【早】"] = (df[cols] == "早").sum(axis=1)
    df["【Q】"] = (df[cols] == "Q").sum(axis=1)
    df["【遅】"] = (df[cols] == "遅").sum(axis=1)
    df["【入】"] = (df[cols] == "入").sum(axis=1)
    df["【明】"] = (df[cols] == "明").sum(axis=1)
    df["【休計】"] = ((df[cols] == "休") | (df[cols] == "希望休")).sum(axis=1)
    df["【有給】"] = (df[cols] == "有給").sum(axis=1)
    
    if (df[cols] == "研修").any().any():
        df["【研修】"] = (df[cols] == "研修").sum(axis=1)
    if (df[cols] == "フリー").any().any():
        df["【フリー】"] = (df[cols] == "フリー").sum(axis=1)
        
    return df

def style_shift_df(df):
    def color_cell(val):
        if val in ["休", "希望休"]:
            return "background-color: #ffcccc; color: #cc0000; font-weight: bold;"
        elif val == "有給":
            return "background-color: #cce5ff; color: #004085; font-weight: bold;"
        return ""
    try:
        return df.style.map(color_cell)
    except AttributeError:
        return df.style.applymap(color_cell)

# =====================================================================
# 事前チェック関数
# =====================================================================
def check_shortage():
    if staffs < 10:
        st.error(f"❌ 現在のスタッフ数（{staffs}人）では、1日に必要な勤務枠（早3・遅3・入2・明2＝計10人）を満たせません。")
        return False

    total_slots = staffs * TARGET_DAYS
    work_slots = (3 + 3 + 2 + 2) * TARGET_DAYS + (TARGET_DAYS - len(Q_SHIFT_DAYS))
    specific_other_slots = sum(1 for (s, d), t in SPECIFIC_SHIFTS.items() if s < staffs and t not in [S_YASU, S_HAYA, S_OSO, S_IRI, S_AKE, S_Q])
    
    available_slots = total_slots - work_slots - specific_other_slots
    required_yasu = staffs * TARGET_YASU
    shortage = required_yasu - available_slots
    
    if shortage > 0:
        st.error(f"❌ 【重大なエラー：枠不足】指定された条件では、全員に「休（希望休含む）」を{TARGET_YASU}回配置することが物理的に不可能です。（必要な休の枠に対して 【 {shortage} 枠 】 足りません）")
        return False
        
    for s in range(staffs):
        p_other = sum(1 for (s_idx, d_idx), t_idx in SPECIFIC_SHIFTS.items() if s_idx == s and t_idx not in [S_YASU, S_HAYA, S_OSO, S_IRI, S_AKE, S_Q])
        p_yasu_req = sum(1 for (s_idx, d_idx), t_idx in SPECIFIC_SHIFTS.items() if s_idx == s and t_idx == S_YASU)
        if p_yasu_req > TARGET_YASU:
            st.error(f"❌ 【個人の希望休超過】{staff_names[s]}さんに「希望休・休」が {p_yasu_req} 回指定されています（設定された休の日数 {TARGET_YASU} 回を超えています）。余分な休みは「有給」などに変更してください。")
            return False
        if TARGET_YASU - (TARGET_DAYS - p_other) > 0:
            st.error(f"❌ 【個人の枠不足】{staff_names[s]}さんの有給・研修等の指定が多すぎるため、「休」を{TARGET_YASU}回配置できません。")
            return False
    return True

# =====================================================================
# 計算ロジック（高速化通常モード ＆ デバッグモード）
# =====================================================================
def run_normal_mode():
    q_days_idx = [d - 1 for d in Q_SHIFT_DAYS]
    personal_limits = get_personal_limits()
    
    def attempt_fast_solve(consecutive_limit, strict_rank):
        model = cp_model.CpModel()
        B = {(s, d, t): model.NewBoolVar(f'B_{s}_{d}_{t}') for s in range(staffs) for d in range(TARGET_DAYS) for t in range(9)}

        for s in range(staffs):
            for d in range(TARGET_DAYS):
                model.AddExactlyOne([B[s, d, t] for t in range(9)])
                req_shift = SPECIFIC_SHIFTS.get((s, d+1))
                if req_shift != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
                if req_shift != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
                if req_shift is not None:
                    model.Add(B[s, d, req_shift] == 1)
            
            model.Add(sum(B[s, d, S_YASU] for d in range(TARGET_DAYS)) == TARGET_YASU)

        for d in range(TARGET_DAYS):
            model.Add(sum(B[s, d, S_HAYA] for s in range(staffs)) == 3)
            model.Add(sum(B[s, d, S_OSO]  for s in range(staffs)) == 3)
            model.Add(sum(B[s, d, S_IRI]  for s in range(staffs)) == 2)
            model.Add(sum(B[s, d, S_AKE]  for s in range(staffs)) == 2)
            if d in q_days_idx:
                model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 0)
            else:
                model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 1)

        for s in range(staffs):
            for d in range(TARGET_DAYS - 1):
                model.AddImplication(B[s, d, S_IRI], B[s, d+1, S_AKE])
                model.AddImplication(B[s, d+1, S_AKE], B[s, d, S_IRI])
                model.AddImplication(B[s, d, S_AKE], B[s, d+1, S_YASU])
                if strict_rank:
                    model.AddImplication(B[s, d, S_Q], B[s, d+1, S_HAYA].Not())
                    model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_HAYA].Not())
                    model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_Q].Not())

        for s in range(staffs):
            prev_c = PREV_CONSECUTIVE.get(s, 0)
            if prev_c > 0:
                rem_days = max(1, consecutive_limit - prev_c)
                if rem_days <= TARGET_DAYS:
                    model.Add(sum(B[s, k, S_YASU] + B[s, k, S_YUKU] for k in range(rem_days)) >= 1)

            for d in range(TARGET_DAYS - consecutive_limit + 1):
                model.Add(sum(B[s, d+k, S_YASU] + B[s, d+k, S_YUKU] for k in range(consecutive_limit)) >= 1)

        for s_idx, s_type, max_c, _ in personal_limits:
            if max_c == 0:
                model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) == 0)
            else:
                model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) <= max_c)

        iri_counts = [model.NewIntVar(0, TARGET_DAYS, f'iri_{s}') for s in range(staffs)]
        for s in range(staffs):
            model.Add(iri_counts[s] == sum(B[s, d, S_IRI] for d in range(TARGET_DAYS)))
        max_iri = model.NewIntVar(0, TARGET_DAYS, 'max_iri')
        model.AddMaxEquality(max_iri, iri_counts)
        model.Minimize(max_iri)

        solver = cp_model.CpSolver()
        solver.parameters.num_search_workers = 8
        solver.parameters.max_time_in_seconds = 15.0
        status = solver.Solve(model)
        return status, solver, B, consecutive_limit, strict_rank

    with st.spinner("高速計算中...（ステップ1: 横5マス休 / 順位制約あり）"):
        status, solver, B, used_limit, used_strict = attempt_fast_solve(5, True)
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        st.warning("⚠️ 横5マス制約では組めませんでした。横6マスに緩和して再計算します...")
        with st.spinner("高速再計算中...（ステップ2: 横6マス休 / 順位制約あり）"):
            status, solver, B, used_limit, used_strict = attempt_fast_solve(6, True)
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        st.warning("⚠️ 横6マスでも組めませんでした。順位制約を解除して再計算します...")
        with st.spinner("高速再計算中...（ステップ3: 横6マス休 / 順位制約緩和）"):
            status, solver, B, used_limit, used_strict = attempt_fast_solve(6, False)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        st.success(f"✅ シフトの作成が完了しました！（適用条件：横{used_limit}マス休 / 順位制約 {'あり' if used_strict else '緩和'}）")
        raw_table = [[shift_str[next(t for t in range(9) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)]
        return format_result_df(raw_table)
    else:
        st.error("❌ 最大まで緩和しましたが、条件が厳しすぎて組めません。自動的に「原因調査モード」で衝突箇所を特定します...")
        return run_debug_mode()

def run_debug_mode():
    model = cp_model.CpModel()
    q_days_idx = [d - 1 for d in Q_SHIFT_DAYS]
    personal_limits = get_personal_limits()
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
        add_soft(model.Add(sum(B[s, d, S_YASU] for d in range(TARGET_DAYS)) == TARGET_YASU), 50, f"【休数エラー】{staff_names[s]}さんの「休＋希望休」が{TARGET_YASU}回ではありません")

    for d in range(TARGET_DAYS):
        add_soft(model.Add(sum(B[s, d, S_HAYA] for s in range(staffs)) == 3), 20, f"【人数エラー】{day_labels[d+1]}の「早番」が3人ではありません")
        add_soft(model.Add(sum(B[s, d, S_OSO]  for s in range(staffs)) == 3), 20, f"【人数エラー】{day_labels[d+1]}の「遅番」が3人ではありません")
        add_soft(model.Add(sum(B[s, d, S_IRI]  for s in range(staffs)) == 2), 20, f"【人数エラー】{day_labels[d+1]}の「入」が2人ではありません")
        add_soft(model.Add(sum(B[s, d, S_AKE]  for s in range(staffs)) == 2), 20, f"【人数エラー】{day_labels[d+1]}の「明」が2人ではありません")
        if d in q_days_idx:
            add_soft(model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 0), 20, f"【人数エラー】{day_labels[d+1]}に「Q」が配置されています")
        else:
            add_soft(model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 1), 20, f"【人数エラー】{day_labels[d+1]}の「Q」が1人ではありません")

    for s in range(staffs):
        for d in range(TARGET_DAYS - 1):
            add_soft(model.Add(B[s, d+1, S_AKE] >= B[s, d, S_IRI]), 50, f"【リレー違反】{staff_names[s]}さんの{day_labels[d+1]}「入」の次が「明」ではない")
            add_soft(model.Add(B[s, d, S_IRI] >= B[s, d+1, S_AKE]), 50, f"【リレー違反】{staff_names[s]}さんの{day_labels[d+2]}「明」の前が「入」ではない")
            add_soft(model.Add(B[s, d+1, S_YASU] >= B[s, d, S_AKE]), 50, f"【リレー違反】{staff_names[s]}さんの{day_labels[d+1]}「明」の次が「休」ではない")
            add_soft(model.Add(B[s, d, S_Q] + B[s, d+1, S_HAYA] <= 1), 10, f"【順位違反】{staff_names[s]}さんの{day_labels[d+1]}「Q」の次に「早」")
            add_soft(model.Add(B[s, d, S_OSO] + B[s, d+1, S_HAYA] <= 1), 10, f"【順位違反】{staff_names[s]}さんの{day_labels[d+1]}「遅」の次に「早」")
            add_soft(model.Add(B[s, d, S_OSO] + B[s, d+1, S_Q] <= 1), 10, f"【順位違反】{staff_names[s]}さんの{day_labels[d+1]}「遅」の次に「Q」")

    for s in range(staffs):
        prev_c = PREV_CONSECUTIVE.get(s, 0)
        if prev_c > 0:
            rem_days = max(1, 5 - prev_c)
            if rem_days <= TARGET_DAYS:
                add_soft(
                    model.Add(sum(B[s, k, S_YASU] + B[s, k, S_YUKU] for k in range(rem_days)) >= 1),
                    15,
                    f"【月初連勤超過】{staff_names[s]}さん（前月{prev_c}連勤）が1日目から{rem_days}日以内に休/有給がありません"
                )
        for d in range(TARGET_DAYS - 4):
            add_soft(model.Add(sum(B[s, d+k, S_YASU] + B[s, d+k, S_YUKU] for k in range(5)) >= 1), 5, f"【連勤超過】{staff_names[s]}さんが{day_labels[d+1]}から連続勤務オーバー(横5マス以内に休/有給なし)")

    for s_idx, s_type, max_c, desc in personal_limits:
        if max_c == 0:
            add_soft(model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) == 0), 20, f"【個人条件違反】{desc}")
        else:
            add_soft(model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) <= max_c), 20, f"【個人条件違反】{desc}")

    model.Minimize(sum(penalties))
    solver = cp_model.CpSolver()
    solver.parameters.num_search_workers = 8
    solver.parameters.max_time_in_seconds = 15.0
    
    with st.spinner("原因調査モードで計算しています..."):
        status = solver.Solve(model)
        
    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        broken_rules = [desc for viol_var, desc in violations_info if solver.Value(viol_var)]
        if not broken_rules:
            st.success("✅ ルール違反は一つもありませんでした！")
        else:
            st.error(f"🚨 以下のルールが衝突しているため、一部を破って強制出力しました（計 {len(broken_rules)} 件の違反）：")
            for rule in broken_rules:
                st.write(f"・ {rule}")
        raw_table = [[shift_str[next(t for t in range(9) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)]
        return format_result_df(raw_table)
    else:
        st.error("希望休などの指定自体が矛盾しています。")
        return pd.DataFrame()

# =====================================================================
# 3. 実行ボタンと結果出力（色分け＆集計列付き）
# =====================================================================
st.subheader("3. シフトの作成と出力")
if st.button("🚀 シフト表を作成する", type="primary", use_container_width=True):
    if check_shortage():
        if "通常作成" in mode:
            result_df = run_normal_mode()
        else:
            result_df = run_debug_mode()
            
        if not result_df.empty:
            styled_result = style_shift_df(result_df)
            st.dataframe(styled_result, use_container_width=True, height=min(620, 65 + staffs * 35))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                styled_result.to_excel(writer, sheet_name='完成シフト表')
                edited_staff_df.to_excel(writer, sheet_name='スタッフ設定', index=False)
                save_req_df.to_excel(writer, sheet_name='希望休入力')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 完成したシフト表をExcelでダウンロード（色分け・回数集計・設定データ同梱）",
                data=excel_data,
                file_name=f"shift_schedule_{target_year}_{target_month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
