import io
import calendar
import datetime
import pandas as pd
import streamlit as st
from ortools.sat.python import cp_model

st.set_page_config(page_title="シフト自動作成システム", layout="wide")
st.title("📅 シフト自動作成システム")

# =====================================================================
# 🔒 パスワードロック機能
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
# ★変更点：S_OJTを新設し、全部で10種類のシフトに拡張しました
S_HAYA, S_Q, S_OSO, S_IRI, S_AKE, S_YASU, S_YUKU, S_KEN, S_FREE, S_OJT = range(10)
shift_str = {
    S_HAYA: '早', S_Q: 'Q', S_OSO: '遅', S_IRI: '入', 
    S_AKE: '明', S_YASU: '休', S_YUKU: '有給', S_KEN: '研修', S_FREE: 'フリー', S_OJT: 'OJT'
}
str_to_shift = {v: k for k, v in shift_str.items()}
str_to_shift['希望休'] = S_YASU

shift_options = ["", "早", "Q", "遅", "入", "明", "希望休", "休", "有給", "研修", "OJT", "フリー"]
limit_options = ["制限なし"] + [str(i) for i in range(0, 32)]
prev_work_options = [str(i) for i in range(0, 6)]
PREV_COL = "前月連勤"
WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]

DEFAULT_STAFF_NAMES = [chr(ord('A') + i) for i in range(16)]
DEFAULT_Q_DAYS = [1, 3, 6, 7, 8, 13, 14, 15, 17, 20, 22, 24, 27, 28, 29]

def get_default_staff_df():
    data = []
    for name in DEFAULT_STAFF_NAMES:
        data.append({"スタッフ名": name, "早の上限": "制限なし", "Qの上限": "制限なし", "遅の上限": "制限なし", "入の上限": "制限なし"})
    return pd.DataFrame(data)

def get_default_ng_pairs_df():
    return pd.DataFrame({"スタッフ1": pd.Series(dtype="str"), "スタッフ2": pd.Series(dtype="str")})

def get_default_personal_lock_df():
    return pd.DataFrame({"スタッフ名": pd.Series(dtype="str"), "開始日": pd.Series(dtype="str"), "終了日": pd.Series(dtype="str")})

def extract_day_num(col_name):
    s = str(col_name).strip()
    if "(" in s: s = s.split("(")[0].strip()
    return int(s) if s.isdigit() else None

# セッション状態の初期化
now = datetime.datetime.now()
if "editor_key" not in st.session_state: st.session_state.editor_key = 0
if "setting_year" not in st.session_state: st.session_state.setting_year = now.year
if "setting_month" not in st.session_state: st.session_state.setting_month = now.month
if "setting_days" not in st.session_state:
    _, auto_d = calendar.monthrange(st.session_state.setting_year, st.session_state.setting_month)
    st.session_state.setting_days = auto_d
if "setting_yasu" not in st.session_state: st.session_state.setting_yasu = 9
if "setting_q_days" not in st.session_state:
    st.session_state.setting_q_days = [d for d in DEFAULT_Q_DAYS if d <= st.session_state.setting_days]
if "exec_mode" not in st.session_state: st.session_state.exec_mode = "通常作成モード（高速・自動3段階緩和）"
if "lock_until_day" not in st.session_state: st.session_state.lock_until_day = 0
if "base_schedule_df" not in st.session_state: st.session_state.base_schedule_df = None

if "staff_df" not in st.session_state: st.session_state.staff_df = get_default_staff_df()
if "ng_pairs_df" not in st.session_state: st.session_state.ng_pairs_df = get_default_ng_pairs_df()
if "personal_lock_df" not in st.session_state: st.session_state.personal_lock_df = get_default_personal_lock_df()
if "req_df" not in st.session_state:
    init_cols = [PREV_COL] + [str(i) for i in range(1, 32)]
    init_req = pd.DataFrame("", index=st.session_state.staff_df["スタッフ名"].tolist(), columns=init_cols)
    init_req[PREV_COL] = "0"
    st.session_state.req_df = init_req

ek = st.session_state.editor_key

# =====================================================================
# サイドバー
# =====================================================================
st.sidebar.header("⚙️ 毎月の基本設定")
col_y, col_m = st.sidebar.columns(2)
new_year = int(col_y.number_input("対象年", min_value=2024, max_value=2035, value=int(st.session_state.setting_year), key=f"ym_year_{ek}"))
new_month = int(col_m.number_input("対象月", min_value=1, max_value=12, value=int(st.session_state.setting_month), key=f"ym_month_{ek}"))

if new_year != st.session_state.setting_year or new_month != st.session_state.setting_month:
    st.session_state.setting_year = new_year
    st.session_state.setting_month = new_month
    _, max_d = calendar.monthrange(new_year, new_month)
    st.session_state.setting_days = max_d
    st.session_state.setting_q_days = [d for d in st.session_state.setting_q_days if d <= max_d]
    st.session_state.editor_key += 1
    st.rerun()

target_year = st.session_state.setting_year
target_month = st.session_state.setting_month
TARGET_DAYS = int(st.sidebar.number_input("月の日数（自動計算）", min_value=28, max_value=31, value=int(st.session_state.setting_days), key=f"days_input_{ek}"))
st.session_state.setting_days = TARGET_DAYS

_, cal_max_days = calendar.monthrange(target_year, target_month)
day_labels = {}
for d in range(1, TARGET_DAYS + 1):
    if d <= cal_max_days: wd = WEEKDAYS_JP[datetime.date(target_year, target_month, d).weekday()]
    else: wd = ""
    day_labels[d] = f"{d}({wd})" if wd else f"{d}"

TARGET_YASU = int(st.sidebar.number_input("全員が必ずとる「休＋希望休」の日数", min_value=1, max_value=15, value=int(st.session_state.setting_yasu), key=f"yasu_input_{ek}"))
st.session_state.setting_yasu = TARGET_YASU
valid_q_defaults = [d for d in st.session_state.setting_q_days if 1 <= d <= TARGET_DAYS]
Q_SHIFT_DAYS = st.sidebar.multiselect("「Q」シフトを【配置しない】日", options=list(range(1, TARGET_DAYS + 1)), default=valid_q_defaults, format_func=lambda d: day_labels.get(d, str(d)), key=f"q_days_input_{ek}_{TARGET_DAYS}")
st.session_state.setting_q_days = Q_SHIFT_DAYS

mode_options = ["通常作成モード（高速・自動3段階緩和）", "🛠️ 欠勤フォロー（最小変更）モード", "原因調査モード（デバッグ用・強制出力）"]
default_mode_idx = mode_options.index(st.session_state.exec_mode) if st.session_state.exec_mode in mode_options else 0
mode = st.sidebar.radio("実行モード", mode_options, index=default_mode_idx, key=f"mode_radio_{ek}")
st.session_state.exec_mode = mode

if "欠勤フォロー" in mode:
    st.sidebar.info("💡 **欠勤フォロー手順**\n1. 下の表で欠勤者の該当日のシフトを書き換えます。\n2. すでに終わった日（動かしたくない過去日）を下で指定して作成ボタンを押します。")
    LOCK_UNTIL_DAY = int(st.sidebar.number_input("🔒 何日目までを過去日として完全固定するか（0＝固定なし）", min_value=0, max_value=TARGET_DAYS - 1, value=min(int(st.session_state.lock_until_day), TARGET_DAYS - 1)))
    st.session_state.lock_until_day = LOCK_UNTIL_DAY
else:
    LOCK_UNTIL_DAY = 0

st.sidebar.divider()
st.sidebar.subheader("📂 保存データの読み込み")

def load_common_settings(xls):
    if "基本設定" in xls.sheet_names:
        loaded_settings = pd.read_excel(xls, "基本設定")
        settings_map = dict(zip(loaded_settings["設定項目"].astype(str), loaded_settings["設定値"]))
        if "対象年" in settings_map and str(settings_map["対象年"]).replace('.', '', 1).isdigit(): st.session_state.setting_year = int(float(settings_map["対象年"]))
        if "対象月" in settings_map and str(settings_map["対象月"]).replace('.', '', 1).isdigit(): st.session_state.setting_month = int(float(settings_map["対象月"]))
        if "月の日数" in settings_map and str(settings_map["月の日数"]).replace('.', '', 1).isdigit(): st.session_state.setting_days = int(float(settings_map["月の日数"]))
        if "休の日数" in settings_map and str(settings_map["休の日数"]).replace('.', '', 1).isdigit(): st.session_state.setting_yasu = int(float(settings_map["休の日数"]))
        if "Qを配置しない日" in settings_map:
            q_raw = str(settings_map["Qを配置しない日"]).strip()
            if q_raw == "" or q_raw.lower() == "nan" or q_raw == "なし": st.session_state.setting_q_days = []
            else: st.session_state.setting_q_days = [int(x.strip().split(".")[0]) for x in q_raw.split(",") if x.strip().split(".")[0].isdigit()]
    if "スタッフ設定" in xls.sheet_names:
        st.session_state.staff_df = pd.read_excel(xls, "スタッフ設定").fillna("制限なし").astype(str)
    if "NGペア設定" in xls.sheet_names:
        loaded_ng = pd.read_excel(xls, "NGペア設定").fillna("").astype(str)
        if "スタッフ1" in loaded_ng.columns and "スタッフ2" in loaded_ng.columns: st.session_state.ng_pairs_df = loaded_ng[["スタッフ1", "スタッフ2"]]
    if "個別期間ロック設定" in xls.sheet_names:
        loaded_plock = pd.read_excel(xls, "個別期間ロック設定").fillna("").astype(str)
        if "スタッフ名" in loaded_plock.columns and "開始日" in loaded_plock.columns and "終了日" in loaded_plock.columns:
            st.session_state.personal_lock_df = loaded_plock[["スタッフ名", "開始日", "終了日"]]

uploaded_file = st.sidebar.file_uploader("Excelファイルを選択", type=["xlsx"])
if uploaded_file is not None:
    if st.sidebar.button("📥 入力途中のデータ（希望休）を復元", use_container_width=True):
        try:
            xls = pd.ExcelFile(uploaded_file)
            load_common_settings(xls)
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
                        if d_num and 1 <= d_num <= 31: normalized_req[str(d_num)] = loaded_req[col].values
                st.session_state.req_df = normalized_req
            st.session_state.base_schedule_df = None
            st.session_state.exec_mode = "通常作成モード（高速・自動3段階緩和）"
            st.session_state.editor_key += 1
            st.sidebar.success("✅ 希望休・設定データを復元しました！")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"読み込みエラー: {e}")

    if st.sidebar.button("🛠️ 完成シフト表を読み込む（欠勤フォロー用）", type="primary", use_container_width=True):
        try:
            xls = pd.ExcelFile(uploaded_file)
            load_common_settings(xls)
            target_sheet = "完成シフト表" if "完成シフト表" in xls.sheet_names else xls.sheet_names[0]
            loaded_sched = pd.read_excel(xls, target_sheet, index_col=0).fillna("").astype(str)
            
            norm_cols = [PREV_COL] + [str(i) for i in range(1, 32)]
            normalized_req = pd.DataFrame("", index=loaded_sched.index.astype(str), columns=norm_cols)
            normalized_req[PREV_COL] = "0"
            
            if "希望休入力" in xls.sheet_names:
                loaded_req = pd.read_excel(xls, "希望休入力", index_col=0).fillna("").astype(str)
                if PREV_COL in loaded_req.columns:
                    for idx_name in normalized_req.index:
                        if idx_name in loaded_req.index:
                            v = str(loaded_req.loc[idx_name, PREV_COL]).strip()
                            if v.replace('.', '', 1).isdigit(): normalized_req.loc[idx_name, PREV_COL] = str(int(float(v)))

            for col in loaded_sched.columns:
                d_num = extract_day_num(col)
                if d_num and 1 <= d_num <= 31: normalized_req[str(d_num)] = [str(v).strip() for v in loaded_sched[col].values]

            if "スタッフ設定" not in xls.sheet_names:
                new_s_list = [{"スタッフ名": s_name, "早の上限": "制限なし", "Qの上限": "制限なし", "遅の上限": "制限なし", "入の上限": "制限なし"} for s_name in loaded_sched.index.astype(str)]
                st.session_state.staff_df = pd.DataFrame(new_s_list)

            st.session_state.req_df = normalized_req.copy()
            st.session_state.base_schedule_df = normalized_req.copy()
            st.session_state.exec_mode = "🛠️ 欠勤フォロー（最小変更）モード"
            st.session_state.editor_key += 1
            st.sidebar.success("✅ 完成シフト表を読み込み、欠勤フォローモードに切り替えました！")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"完成シフト表の読み込みエラー: {e}")

# =====================================================================
# 1. スタッフ名・個人条件・NGペア等の設定
# =====================================================================
st.subheader("1. スタッフ名・個人条件の設定")
with st.expander("👥 スタッフ名（個人名への書き換え）・シフト上限・禁止条件を編集する", expanded=False):
    with st.popover("📝 スタッフ名をまとめて一括変更する（コピペ入力）"):
        bulk_names_text = st.text_area("スタッフ名リストを入力（例：佐藤、鈴木、高橋...）", height=150)
        if st.button("名前を一括反映する", type="primary"):
            raw_list = [n.strip() for n in bulk_names_text.replace("、", "\n").replace(",", "\n").replace(" ", "\n").replace(" ", "\n").splitlines() if n.strip()]
            if raw_list:
                new_staff_df = st.session_state.staff_df.copy()
                old_req_df = st.session_state.req_df.copy()
                for idx, new_name in enumerate(raw_list):
                    if idx < len(new_staff_df): new_staff_df.loc[idx, "スタッフ名"] = new_name
                    else: new_staff_df = pd.concat([new_staff_df, pd.DataFrame([{"スタッフ名": new_name, "早の上限": "制限なし", "Qの上限": "制限なし", "遅の上限": "制限なし", "入の上限": "制限なし"}])], ignore_index=True)
                st.session_state.staff_df = new_staff_df
                updated_names = new_staff_df["スタッフ名"].tolist()
                new_req_df = pd.DataFrame("", index=updated_names, columns=old_req_df.columns)
                new_req_df[PREV_COL] = "0"
                for i in range(min(len(old_req_df), len(updated_names))): new_req_df.iloc[i] = old_req_df.iloc[i].values
                st.session_state.req_df = new_req_df
                st.session_state.editor_key += 1
                st.rerun()

    staff_col_config = {
        "スタッフ名": st.column_config.TextColumn("スタッフ名", required=True, width="medium"),
        "早の上限": st.column_config.SelectboxColumn("早の上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "Qの上限": st.column_config.SelectboxColumn("Qの上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "遅の上限": st.column_config.SelectboxColumn("遅の上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "入の上限": st.column_config.SelectboxColumn("入の上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
    }
    edited_staff_df = st.data_editor(st.session_state.staff_df, column_config=staff_col_config, num_rows="dynamic", use_container_width=True, height=min(420, 65 + len(st.session_state.staff_df) * 35), key=f"staff_editor_{ek}")

raw_names = [str(n).strip() if pd.notna(n) and str(n).strip() != "" else f"スタッフ{i+1}" for i, n in enumerate(edited_staff_df["スタッフ名"])]
staff_names = []
for i, name in enumerate(raw_names):
    if name in staff_names: staff_names.append(f"{name}_{i+1}")
    else: staff_names.append(name)
staffs = len(staff_names)

with st.expander("🚫 夜勤（入）の同日NGペア設定", expanded=False):
    ng_col_config = {
        "スタッフ1": st.column_config.SelectboxColumn("スタッフ1", options=staff_names, required=True, width="medium"),
        "スタッフ2": st.column_config.SelectboxColumn("スタッフ2", options=staff_names, required=True, width="medium"),
    }
    edited_ng_pairs_df = st.data_editor(st.session_state.ng_pairs_df, column_config=ng_col_config, num_rows="dynamic", use_container_width=True, key=f"ng_editor_{ek}")
st.session_state.ng_pairs_df = edited_ng_pairs_df

with st.expander("🔒 欠勤フォロー用：個別期間ロック設定（出勤停止などの期間を完全固定）", expanded=False):
    st.caption("・この期間内はAIによる自動調整が一切行われず、画面上で人間が指定した配分（有給や休）が100%守られます。")
    plock_col_config = {
        "スタッフ名": st.column_config.SelectboxColumn("スタッフ名", options=staff_names, required=True, width="medium"),
        "開始日": st.column_config.SelectboxColumn("開始日（〇日）", options=[str(i) for i in range(1, 32)], required=True, width="small"),
        "終了日": st.column_config.SelectboxColumn("終了日（〇日）", options=[str(i) for i in range(1, 32)], required=True, width="small"),
    }
    edited_personal_lock_df = st.data_editor(st.session_state.personal_lock_df, column_config=plock_col_config, num_rows="dynamic", use_container_width=True, key=f"plock_editor_{ek}")
st.session_state.personal_lock_df = edited_personal_lock_df

# =====================================================================
# 2. メイン画面：表
# =====================================================================
if "欠勤フォロー" in mode:
    st.subheader(f"2. 【欠勤フォローモード】現在のシフト表と欠勤箇所の編集（{target_year}年{target_month}月）")
    if LOCK_UNTIL_DAY > 0:
        st.info(f"🔒 **1日 〜 {day_labels[LOCK_UNTIL_DAY]} は固定されます。**\n・出勤停止などの場合は「個別期間ロック設定」も併用し、対象日のマスを人間が管理したい配分（休・有給）に書き換えてください。")
else:
    st.subheader(f"2. 希望休・指定シフトの入力（{target_year}年{target_month}月）")

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

column_configuration = {PREV_COL: st.column_config.SelectboxColumn("前月連勤", options=prev_work_options, default="0", width="small", required=True)}
for i in range(1, TARGET_DAYS + 1):
    col_label = f"🔒{day_labels[i]}" if ("欠勤フォロー" in mode and i <= LOCK_UNTIL_DAY) else day_labels[i]
    column_configuration[str(i)] = st.column_config.SelectboxColumn(label=col_label, options=shift_options, default="", width="small")

edited_req_df = st.data_editor(current_req_df, column_config=column_configuration, use_container_width=True, height=min(620, 65 + staffs * 35), key=f"req_editor_{ek}")
st.session_state.staff_df = edited_staff_df
st.session_state.req_df = edited_req_df
if "欠勤フォロー" in mode and st.session_state.base_schedule_df is None: st.session_state.base_schedule_df = edited_req_df.copy()

settings_df = pd.DataFrame([
    {"設定項目": "対象年", "設定値": str(target_year)}, {"設定項目": "対象月", "設定値": str(target_month)},
    {"設定項目": "月の日数", "設定値": str(TARGET_DAYS)}, {"設定項目": "休の日数", "設定値": str(TARGET_YASU)},
    {"設定項目": "Qを配置しない日", "設定値": ",".join(map(str, sorted(Q_SHIFT_DAYS))) if Q_SHIFT_DAYS else "なし"},
])
save_req_df = edited_req_df.copy()
save_req_df.columns = [PREV_COL] + [day_labels[int(c)] for c in day_cols]
save_output = io.BytesIO()
with pd.ExcelWriter(save_output, engine='openpyxl') as writer:
    settings_df.to_excel(writer, sheet_name='基本設定', index=False)
    edited_staff_df.to_excel(writer, sheet_name='スタッフ設定', index=False)
    edited_ng_pairs_df.to_excel(writer, sheet_name='NGペア設定', index=False)
    edited_personal_lock_df.to_excel(writer, sheet_name='個別期間ロック設定', index=False)
    save_req_df.to_excel(writer, sheet_name='希望休入力')
save_excel_bytes = save_output.getvalue()

col_save1, col_save2 = st.columns([1, 1])
with col_save1: st.download_button("💾 現在の入力進捗をExcelに一時保存", data=save_excel_bytes, file_name=f"shift_progress_{target_year}_{target_month:02d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
with col_save2:
    if "欠勤フォロー" in mode:
        if st.button("🧹 表を空欄に戻して「通常作成モード」へ戻る", use_container_width=True):
            empty_req = pd.DataFrame("", index=staff_names, columns=[PREV_COL] + day_cols)
            empty_req[PREV_COL] = edited_req_df[PREV_COL].values
            st.session_state.req_df = empty_req
            st.session_state.base_schedule_df = None
            st.session_state.exec_mode = "通常作成モード（高速・自動3段階緩和）"
            st.session_state.editor_key += 1
            st.rerun()

st.sidebar.download_button("💾 現在の入力進捗を保存", data=save_excel_bytes, file_name=f"shift_progress_{target_year}_{target_month:02d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
st.sidebar.divider()
if st.sidebar.button("🔓 ログアウト"):
    st.session_state.authenticated = False
    st.rerun()

# =====================================================================
# 画面データ読み取り
# =====================================================================
def get_personal_locks():
    locks = []
    for _, row in edited_personal_lock_df.iterrows():
        s_name = str(row.get("スタッフ名", "")).strip()
        start = str(row.get("開始日", "")).strip()
        end = str(row.get("終了日", "")).strip()
        if s_name in staff_names and start.isdigit() and end.isdigit():
            st_d = min(int(start), int(end))
            en_d = max(int(start), int(end))
            locks.append((staff_names.index(s_name), st_d, en_d))
    return locks

PERSONAL_LOCKS = get_personal_locks()
SPECIFIC_SHIFTS, KIBOU_YASU_CELLS, PREV_CONSECUTIVE, ORIGINAL_SHIFTS, ORIGINAL_STR_SHIFTS = {}, set(), {}, {}, {}
base_df = st.session_state.base_schedule_df

for s_idx, s_name in enumerate(staff_names):
    p_val = str(edited_req_df.loc[s_name, PREV_COL]).strip()
    PREV_CONSECUTIVE[s_idx] = int(p_val) if p_val.isdigit() else 0
    for d in range(1, TARGET_DAYS + 1):
        val_str = str(edited_req_df.loc[s_name, str(d)]).strip() if pd.notna(edited_req_df.loc[s_name, str(d)]) else ""
        if val_str == "希望休": KIBOU_YASU_CELLS.add((s_idx, d - 1))

        orig_str = ""
        if base_df is not None:
            if s_name in base_df.index and str(d) in base_df.columns: orig_str = str(base_df.loc[s_name, str(d)]).strip()
            elif s_idx < len(base_df.index) and str(d) in base_df.columns: orig_str = str(base_df.iloc[s_idx][str(d)]).strip()
        if orig_str in str_to_shift:
            ORIGINAL_SHIFTS[(s_idx, d)] = str_to_shift[orig_str]
            ORIGINAL_STR_SHIFTS[(s_idx, d)] = orig_str

        is_personal_locked = any(s == s_idx and start <= d <= end for s, start, end in PERSONAL_LOCKS)

        if "欠勤フォロー" not in mode:
            if val_str in str_to_shift: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]
        else:
            if val_str in str_to_shift:
                # OJTとフリーは「絶対固定（ハードロック）」の対象から外し、AIが自由に変更（バッファ利用）できるようにする
                if d <= LOCK_UNTIL_DAY: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]
                elif is_personal_locked: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]
                elif orig_str != "" and val_str != orig_str: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]
                elif val_str in ["希望休", "休", "有給", "研修"]: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]

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
                limits.append((s_idx, s_type, max_c, f"{s_name}さんの「{s_label}」は禁止(0回)" if max_c == 0 else f"{s_name}さんの「{s_label}」は{max_c}回まで"))
    return limits

def get_ng_pairs():
    pairs = []
    seen = set()
    for _, row in edited_ng_pairs_df.iterrows():
        n1, n2 = str(row.get("スタッフ1", "")).strip(), str(row.get("スタッフ2", "")).strip()
        if n1 in staff_names and n2 in staff_names and n1 != n2:
            idx1, idx2 = staff_names.index(n1), staff_names.index(n2)
            pair_key = tuple(sorted((idx1, idx2)))
            if pair_key not in seen:
                seen.add(pair_key)
                pairs.append((idx1, idx2, f"{n1}さんと{n2}さんの同日「入」禁止"))
    return pairs

def format_result_df(raw_table_data):
    for s_idx, d_idx in KIBOU_YASU_CELLS:
        if 0 <= s_idx < len(raw_table_data) and 0 <= d_idx < TARGET_DAYS:
            if raw_table_data[s_idx][d_idx] == '休': raw_table_data[s_idx][d_idx] = '希望休'
    cols = [day_labels[i] for i in range(1, TARGET_DAYS + 1)]
    df = pd.DataFrame(raw_table_data, index=staff_names, columns=cols)
    df["【早】"], df["【Q】"], df["【遅】"], df["【入】"], df["【明】"] = (df[cols] == "早").sum(axis=1), (df[cols] == "Q").sum(axis=1), (df[cols] == "遅").sum(axis=1), (df[cols] == "入").sum(axis=1), (df[cols] == "明").sum(axis=1)
    df["【休計】"], df["【有給】"] = ((df[cols] == "休") | (df[cols] == "希望休")).sum(axis=1), (df[cols] == "有給").sum(axis=1)
    if (df[cols] == "研修").any().any(): df["【研修】"] = (df[cols] == "研修").sum(axis=1)
    if (df[cols] == "OJT").any().any(): df["【OJT】"] = (df[cols] == "OJT").sum(axis=1)
    if (df[cols] == "フリー").any().any(): df["【フリー】"] = (df[cols] == "フリー").sum(axis=1)
    return df

def style_shift_df(df):
    def color_cell(val):
        if val in ["休", "希望休"]: return "background-color: #ffcccc; color: #cc0000; font-weight: bold;"
        elif val == "有給": return "background-color: #cce5ff; color: #004085; font-weight: bold;"
        return ""
    try: return df.style.map(color_cell)
    except AttributeError: return df.style.applymap(color_cell)

def check_shortage():
    if staffs < 10:
        st.error(f"❌ 現在のスタッフ数（{staffs}人）では、1日に必要な勤務枠を満たせません。")
        return False
    if "欠勤フォロー" in mode: return True
    total_slots = staffs * TARGET_DAYS
    work_slots = (3 + 3 + 2 + 2) * TARGET_DAYS + (TARGET_DAYS - len(Q_SHIFT_DAYS))
    specific_other_slots = sum(1 for (s, d), t in SPECIFIC_SHIFTS.items() if s < staffs and t not in [S_YASU, S_HAYA, S_OSO, S_IRI, S_AKE, S_Q])
    if (staffs * TARGET_YASU) - (total_slots - work_slots - specific_other_slots) > 0:
        st.error(f"❌ 【枠不足】全員に「休」を{TARGET_YASU}回配置することが物理的に不可能です。")
        return False
    for s in range(staffs):
        if sum(1 for (s_idx, d_idx), t_idx in SPECIFIC_SHIFTS.items() if s_idx == s and t_idx == S_YASU) > TARGET_YASU:
            st.error(f"❌ 【希望休超過】{staff_names[s]}さんに休が {TARGET_YASU} 回を超えて指定されています。")
            return False
    return True

# =====================================================================
# 計算ロジック1：通常モード
# =====================================================================
def run_normal_mode():
    q_days_idx = [d - 1 for d in Q_SHIFT_DAYS]
    personal_limits = get_personal_limits()
    ng_pairs = get_ng_pairs()
    
    def attempt_fast_solve(consecutive_limit, strict_rank):
        model = cp_model.CpModel()
        # ★変数を10種類(S_OJT追加)に対応
        B = {(s, d, t): model.NewBoolVar(f'B_{s}_{d}_{t}') for s in range(staffs) for d in range(TARGET_DAYS) for t in range(10)}

        for s in range(staffs):
            for d in range(TARGET_DAYS):
                model.AddExactlyOne([B[s, d, t] for t in range(10)])
                req_shift = SPECIFIC_SHIFTS.get((s, d+1))
                if req_shift != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
                if req_shift != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
                if req_shift != S_FREE: model.Add(B[s, d, S_FREE] == 0)
                if req_shift != S_OJT:  model.Add(B[s, d, S_OJT] == 0)
                if req_shift is not None: model.Add(B[s, d, req_shift] == 1)
            model.Add(sum(B[s, d, S_YASU] for d in range(TARGET_DAYS)) == TARGET_YASU)

        for d in range(TARGET_DAYS):
            model.Add(sum(B[s, d, S_HAYA] for s in range(staffs)) == 3)
            model.Add(sum(B[s, d, S_OSO]  for s in range(staffs)) == 3)
            model.Add(sum(B[s, d, S_IRI]  for s in range(staffs)) == 2)
            model.Add(sum(B[s, d, S_AKE]  for s in range(staffs)) == 2)
            if d in q_days_idx: model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 0)
            else: model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 1)

        for s in range(staffs):
            for d in range(TARGET_DAYS - 1):
                model.AddImplication(B[s, d, S_IRI], B[s, d+1, S_AKE])
                model.AddImplication(B[s, d+1, S_AKE], B[s, d, S_IRI])
                model.Add(B[s, d+1, S_YASU] + B[s, d+1, S_YUKU] >= B[s, d, S_AKE])
                if strict_rank:
                    model.AddImplication(B[s, d, S_Q], B[s, d+1, S_HAYA].Not())
                    model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_HAYA].Not())
                    model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_Q].Not())

        for s in range(staffs):
            prev_c = PREV_CONSECUTIVE.get(s, 0)
            if prev_c > 0:
                rem_days = max(1, consecutive_limit - prev_c)
                if rem_days <= TARGET_DAYS: model.Add(sum(B[s, k, S_YASU] + B[s, k, S_YUKU] for k in range(rem_days)) >= 1)
            for d in range(TARGET_DAYS - consecutive_limit + 1):
                model.Add(sum(B[s, d+k, S_YASU] + B[s, d+k, S_YUKU] for k in range(consecutive_limit)) >= 1)

        for s_idx, s_type, max_c, _ in personal_limits:
            if max_c == 0: model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) == 0)
            else: model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) <= max_c)

        for s1, s2, _ in ng_pairs:
            for d in range(TARGET_DAYS): model.Add(B[s1, d, S_IRI] + B[s2, d, S_IRI] <= 1)

        iri_counts = [model.NewIntVar(0, TARGET_DAYS, f'iri_{s}') for s in range(staffs)]
        for s in range(staffs): model.Add(iri_counts[s] == sum(B[s, d, S_IRI] for d in range(TARGET_DAYS)))
        max_iri = model.NewIntVar(0, TARGET_DAYS, 'max_iri')
        model.AddMaxEquality(max_iri, iri_counts)
        model.Minimize(max_iri)

        solver = cp_model.CpSolver()
        solver.parameters.num_search_workers = 8
        solver.parameters.max_time_in_seconds = 15.0
        return solver.Solve(model), solver, B, consecutive_limit, strict_rank

    with st.spinner("高速計算中...（ステップ1）"): status, solver, B, used_limit, used_strict = attempt_fast_solve(5, True)
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        with st.spinner("高速再計算中...（ステップ2）"): status, solver, B, used_limit, used_strict = attempt_fast_solve(6, True)
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        with st.spinner("高速再計算中...（ステップ3）"): status, solver, B, used_limit, used_strict = attempt_fast_solve(6, False)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        st.success("✅ シフト作成完了！")
        res_df = format_result_df([[shift_str[next(t for t in range(10) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)])
        return res_df, pd.DataFrame()
    else:
        st.error("❌ 条件が厳しすぎます。原因調査モードへ移行します...")
        return run_debug_mode()

# =====================================================================
# 計算ロジック2：🛠️ 欠勤フォロー（要件定義準拠・完全ルール）
# =====================================================================
def run_reschedule_mode():
    q_days_idx = [d - 1 for d in Q_SHIFT_DAYS]
    personal_limits = get_personal_limits()
    ng_pairs = get_ng_pairs()
    start_check_d = max(0, LOCK_UNTIL_DAY - 1)

    def attempt_reschedule(consecutive_limit, strict_rank):
        model = cp_model.CpModel()
        B = {(s, d, t): model.NewBoolVar(f'B_{s}_{d}_{t}') for s in range(staffs) for d in range(TARGET_DAYS) for t in range(10)}
        obj_terms = []

        is_modified = [model.NewBoolVar(f'is_mod_{s}') for s in range(staffs)]
        for s in range(staffs):
            # ③ 関わる人数を最小限にする（ペナルティ1000）
            obj_terms.append(is_modified[s] * 1000)

        for s in range(staffs):
            # ① 公休数を規定数丁度にすること（絶対条件）
            model.Add(sum(B[s, d, S_YASU] for d in range(TARGET_DAYS)) == TARGET_YASU)

            for d in range(TARGET_DAYS):
                model.AddExactlyOne([B[s, d, t] for t in range(10)])
                
                req_shift = SPECIFIC_SHIFTS.get((s, d+1))
                orig_t = ORIGINAL_SHIFTS.get((s, d+1))
                is_p_locked = any(p_s == s and start <= (d+1) <= end for p_s, start, end in PERSONAL_LOCKS)

                # ベース制限：AIが勝手に有給・研修・フリー・OJTを新規作成するのを防ぐ
                if req_shift != S_YUKU and orig_t != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
                if req_shift != S_KEN and orig_t != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
                if req_shift != S_FREE and orig_t != S_FREE: model.Add(B[s, d, S_FREE] == 0)
                if req_shift != S_OJT and orig_t != S_OJT: model.Add(B[s, d, S_OJT] == 0)

                # 期間ロックや過去日は絶対固定
                if (d + 1) <= LOCK_UNTIL_DAY or is_p_locked:
                    if req_shift is not None:
                        model.Add(B[s, d, req_shift] == 1)
                    elif orig_t is not None:
                        model.Add(B[s, d, orig_t] == 1)
                else:
                    if req_shift == S_YUKU:
                        # ① 例外：有給を休に変更することは許可（ペナルティは最も軽い1）
                        model.Add(B[s, d, S_YUKU] + B[s, d, S_YASU] == 1)
                        obj_terms.append(B[s, d, S_YASU] * 1)
                    elif req_shift is not None:
                        # ① 指定された希望休・休・出勤・研修などは絶対に固定
                        model.Add(B[s, d, req_shift] == 1)
                        if orig_t is not None and req_shift != orig_t:
                            # ユーザーが意図的に変更したマスのため、この人は「変更に関わった人」とみなす
                            model.Add(is_modified[s] == 1)

                    if orig_t is not None:
                        if req_shift == S_YUKU and orig_t == S_YUKU:
                            # 有給から休への変換は「人を巻き込んだ」としてカウントしない
                            pass
                        elif req_shift is None:
                            # AIが自由に動かせるマス（OJT・フリー・出勤等）
                            if orig_t in [S_OJT, S_FREE]:
                                # OJTとフリーはバッファ枠のため超低ペナルティ（5）で変更可能、かつ人を巻き込んだカウントにも入れない
                                obj_terms.append((1 - B[s, d, orig_t]) * 5)
                            else:
                                if orig_t == S_YASU:
                                    # 元々休だった日を動かすのはペナルティ20
                                    obj_terms.append((1 - B[s, d, S_YASU]) * 20)
                                else:
                                    # ④ 関わるコマ数を最小にする（出勤の移動はペナルティ10）
                                    obj_terms.append((1 - B[s, d, orig_t]) * 10)
                                # 元のシフトから変更されたら「巻き込まれた人」としてカウントする
                                model.AddImplication(B[s, d, orig_t].Not(), is_modified[s])

        # ② 一日の必要要件（早や遅の規定数）を守る事
        for d in range(LOCK_UNTIL_DAY, TARGET_DAYS):
            model.Add(sum(B[s, d, S_HAYA] for s in range(staffs)) == 3)
            model.Add(sum(B[s, d, S_OSO]  for s in range(staffs)) == 3)
            model.Add(sum(B[s, d, S_IRI]  for s in range(staffs)) == 2)
            model.Add(sum(B[s, d, S_AKE]  for s in range(staffs)) == 2)
            if d in q_days_idx: model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 0)
            else: model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 1)

        # 順序ルール
        for s in range(staffs):
            for d in range(start_check_d, TARGET_DAYS - 1):
                model.AddImplication(B[s, d, S_IRI], B[s, d+1, S_AKE])
                if d + 1 >= LOCK_UNTIL_DAY: model.AddImplication(B[s, d+1, S_AKE], B[s, d, S_IRI])
                model.Add(B[s, d+1, S_YASU] + B[s, d+1, S_YUKU] >= B[s, d, S_AKE])
                if strict_rank:
                    model.AddImplication(B[s, d, S_Q], B[s, d+1, S_HAYA].Not())
                    model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_HAYA].Not())
                    model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_Q].Not())

        # 連勤ルール
        for s in range(staffs):
            if LOCK_UNTIL_DAY == 0:
                prev_c = PREV_CONSECUTIVE.get(s, 0)
                if prev_c > 0:
                    rem_days = max(1, consecutive_limit - prev_c)
                    if rem_days <= TARGET_DAYS: model.Add(sum(B[s, k, S_YASU] + B[s, k, S_YUKU] for k in range(rem_days)) >= 1)
            for d in range(TARGET_DAYS - consecutive_limit + 1):
                if d + consecutive_limit - 1 >= LOCK_UNTIL_DAY:
                    model.Add(sum(B[s, d+k, S_YASU] + B[s, d+k, S_YUKU] for k in range(consecutive_limit)) >= 1)

        # 個人上限ルール
        for s_idx, s_type, max_c, _ in personal_limits:
            locked_cnt = sum(1 for d in range(1, LOCK_UNTIL_DAY + 1) if SPECIFIC_SHIFTS.get((s_idx, d)) == s_type)
            eff_max = max(max_c, locked_cnt)
            if eff_max == 0: model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) == 0)
            else: model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) <= eff_max)

        # NGペア
        for s1, s2, _ in ng_pairs:
            for d in range(LOCK_UNTIL_DAY, TARGET_DAYS): model.Add(B[s1, d, S_IRI] + B[s2, d, S_IRI] <= 1)

        model.Minimize(sum(obj_terms))
        solver = cp_model.CpSolver()
        solver.parameters.num_search_workers = 8
        solver.parameters.max_time_in_seconds = 15.0
        return solver.Solve(model), solver, B, consecutive_limit, strict_rank

    steps = [
        (5, True, "ステップ1"), (6, True, "ステップ2"), (6, False, "ステップ3")
    ]
    for c_lim, s_rank, label in steps:
        with st.spinner(f"最小変更で再計算中...（{label}）"):
            status, solver, B, used_limit, used_strict = attempt_reschedule(c_lim, s_rank)
        if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]: break

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        raw_table = [[shift_str[next(t for t in range(10) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)]
        result_df = format_result_df(raw_table)

        ui_changes = []
        change_records = []
        for s_idx, s_name in enumerate(staff_names):
            for d in range(1, TARGET_DAYS + 1):
                orig_s = ORIGINAL_STR_SHIFTS.get((s_idx, d), "")
                new_s = result_df.loc[s_name, day_labels[d]]
                if orig_s != "" and orig_s != new_s:
                    tag = "手動変更（欠勤・指定）" if (s_idx, d) in SPECIFIC_SHIFTS and d > LOCK_UNTIL_DAY else "自動調整（フォロー）"
                    ui_changes.append(f"・ **{day_labels[d]} {s_name}さん** ： `{orig_s}` ➡️ **`{new_s}`**  【{tag}】")
                    change_records.append({
                        "日付": day_labels[d],
                        "スタッフ名": s_name,
                        "変更前": orig_s,
                        "変更後": new_s,
                        "備考": tag
                    })

        changes_df = pd.DataFrame(change_records)
        st.success(f"✅ 最小変更でのシフト再作成が完了しました！（総変更：{len(ui_changes)} 箇所）")
        with st.expander(f"📝 元のシフト表から変更された箇所の一覧を見る（計 {len(ui_changes)} 箇所）", expanded=True):
            if not ui_changes: st.write("変更された箇所はありません。")
            else:
                for ch in ui_changes: st.markdown(ch)
        return result_df, changes_df
    else:
        st.error("❌ 最小変更モードでは組めませんでした。原因調査モードへ移行します...")
        return run_debug_mode()

# =====================================================================
# 計算ロジック3：原因調査モード
# =====================================================================
def run_debug_mode():
    model = cp_model.CpModel()
    q_days_idx = [d - 1 for d in Q_SHIFT_DAYS]
    personal_limits = get_personal_limits()
    ng_pairs = get_ng_pairs()
    B = {(s, d, t): model.NewBoolVar(f'B_{s}_{d}_{t}') for s in range(staffs) for d in range(TARGET_DAYS) for t in range(10)}
    penalties = []
    violations_info = []

    def add_soft(constraint, weight, desc):
        viol = model.NewBoolVar('')
        constraint.OnlyEnforceIf(viol.Not())
        penalties.append(viol * weight)
        violations_info.append((viol, desc))

    for s in range(staffs):
        for d in range(TARGET_DAYS):
            model.AddExactlyOne([B[s, d, t] for t in range(10)])
            req_shift = SPECIFIC_SHIFTS.get((s, d+1))
            orig_t = ORIGINAL_SHIFTS.get((s, d+1))
            
            if "欠勤フォロー" not in mode:
                if req_shift != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
                if req_shift != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
                if req_shift != S_FREE: model.Add(B[s, d, S_FREE] == 0)
                if req_shift != S_OJT:  model.Add(B[s, d, S_OJT] == 0)
            else:
                if req_shift != S_YUKU and orig_t != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
                if req_shift != S_KEN and orig_t != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
                if req_shift != S_FREE and orig_t != S_FREE: model.Add(B[s, d, S_FREE] == 0)
                if req_shift != S_OJT and orig_t != S_OJT: model.Add(B[s, d, S_OJT] == 0)
                
            if req_shift is not None: model.Add(B[s, d, req_shift] == 1)

    for s in range(staffs): add_soft(model.Add(sum(B[s, d, S_YASU] for d in range(TARGET_DAYS)) == TARGET_YASU), 50, f"【休数エラー】{staff_names[s]}さんの休が{TARGET_YASU}回ではありません")

    for d in range(TARGET_DAYS):
        add_soft(model.Add(sum(B[s, d, S_HAYA] for s in range(staffs)) == 3), 20, f"【人数エラー】{day_labels[d+1]}の早番が3人ではありません")
        add_soft(model.Add(sum(B[s, d, S_OSO]  for s in range(staffs)) == 3), 20, f"【人数エラー】{day_labels[d+1]}の遅番が3人ではありません")
        add_soft(model.Add(sum(B[s, d, S_IRI]  for s in range(staffs)) == 2), 20, f"【人数エラー】{day_labels[d+1]}の入が2人ではありません")
        add_soft(model.Add(sum(B[s, d, S_AKE]  for s in range(staffs)) == 2), 20, f"【人数エラー】{day_labels[d+1]}の明が2人ではありません")
        if d in q_days_idx: add_soft(model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 0), 20, f"【人数エラー】{day_labels[d+1]}にQが配置されています")
        else: add_soft(model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 1), 20, f"【人数エラー】{day_labels[d+1]}のQが1人ではありません")

    for s in range(staffs):
        for d in range(TARGET_DAYS - 1):
            add_soft(model.Add(B[s, d+1, S_AKE] >= B[s, d, S_IRI]), 50, f"【リレー違反】{staff_names[s]}さんの{day_labels[d+1]}「入」の次が「明」ではない")
            add_soft(model.Add(B[s, d, S_IRI] >= B[s, d+1, S_AKE]), 50, f"【リレー違反】{staff_names[s]}さんの{day_labels[d+2]}「明」の前が「入」ではない")
            add_soft(model.Add(B[s, d+1, S_YASU] + B[s, d+1, S_YUKU] >= B[s, d, S_AKE]), 50, f"【リレー違反】{staff_names[s]}さんの{day_labels[d+1]}「明」の次が休・有給ではない")
            add_soft(model.Add(B[s, d, S_Q] + B[s, d+1, S_HAYA] <= 1), 10, f"【順位違反】{staff_names[s]}さんの{day_labels[d+1]}「Q」の次に「早」")
            add_soft(model.Add(B[s, d, S_OSO] + B[s, d+1, S_HAYA] <= 1), 10, f"【順位違反】{staff_names[s]}さんの{day_labels[d+1]}「遅」の次に「早」")
            add_soft(model.Add(B[s, d, S_OSO] + B[s, d+1, S_Q] <= 1), 10, f"【順位違反】{staff_names[s]}さんの{day_labels[d+1]}「遅」の次に「Q」")

    for s in range(staffs):
        prev_c = PREV_CONSECUTIVE.get(s, 0)
        if prev_c > 0:
            rem_days = max(1, 5 - prev_c)
            if rem_days <= TARGET_DAYS: add_soft(model.Add(sum(B[s, k, S_YASU] + B[s, k, S_YUKU] for k in range(rem_days)) >= 1), 15, f"【月初連勤超過】{staff_names[s]}さんが{rem_days}日以内に休/有給なし")
        for d in range(TARGET_DAYS - 4): add_soft(model.Add(sum(B[s, d+k, S_YASU] + B[s, d+k, S_YUKU] for k in range(5)) >= 1), 5, f"【連勤超過】{staff_names[s]}さんが{day_labels[d+1]}から連続オーバー")

    for s_idx, s_type, max_c, desc in personal_limits:
        if max_c == 0: add_soft(model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) == 0), 20, f"【条件違反】{desc}")
        else: add_soft(model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) <= max_c), 20, f"【条件違反】{desc}")

    for s1, s2, desc in ng_pairs:
        for d in range(TARGET_DAYS): add_soft(model.Add(B[s1, d, S_IRI] + B[s2, d, S_IRI] <= 1), 25, f"【NGペア違反】{day_labels[d+1]}: {desc}")

    model.Minimize(sum(penalties))
    solver = cp_model.CpSolver()
    solver.parameters.num_search_workers = 8
    solver.parameters.max_time_in_seconds = 15.0
    
    with st.spinner("原因調査モードで計算しています..."): status = solver.Solve(model)
        
    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        broken_rules = [desc for viol_var, desc in violations_info if solver.Value(viol_var)]
        if not broken_rules: st.success("✅ ルール違反なし！")
        else:
            st.error(f"🚨 以下のルールが衝突しています（計 {len(broken_rules)} 件）：")
            for rule in broken_rules: st.write(f"・ {rule}")
        raw_table = [[shift_str[next(t for t in range(10) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)]
        return format_result_df(raw_table), pd.DataFrame()
    else:
        st.error("指定自体が矛盾しています。")
        return pd.DataFrame(), pd.DataFrame()

# =====================================================================
# 3. 実行ボタンと結果出力
# =====================================================================
st.subheader("3. シフトの作成と出力")
btn_label = "🛠️ 最小変更でシフトを再作成する（欠勤フォロー実行）" if "欠勤フォロー" in mode else "🚀 シフト表を作成する"

if st.button(btn_label, type="primary", use_container_width=True):
    if check_shortage():
        if "通常作成" in mode: result_df, changes_df = run_normal_mode()
        elif "欠勤フォロー" in mode: result_df, changes_df = run_reschedule_mode()
        else: result_df, changes_df = run_debug_mode()
            
        if not result_df.empty:
            styled_result = style_shift_df(result_df)
            st.dataframe(styled_result, use_container_width=True, height=min(620, 65 + staffs * 35))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                styled_result.to_excel(writer, sheet_name='完成シフト表')
                if not changes_df.empty:
                    changes_df.to_excel(writer, sheet_name='変更依頼リスト', index=False)
                settings_df.to_excel(writer, sheet_name='基本設定', index=False)
                edited_staff_df.to_excel(writer, sheet_name='スタッフ設定', index=False)
                edited_ng_pairs_df.to_excel(writer, sheet_name='NGペア設定', index=False)
                edited_personal_lock_df.to_excel(writer, sheet_name='個別期間ロック設定', index=False)
                save_req_df.to_excel(writer, sheet_name='希望休入力')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 完成したシフト表をExcelでダウンロード（色分け・回数集計・全設定データ同梱）",
                data=excel_data,
                file_name=f"shift_schedule_{target_year}_{target_month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
