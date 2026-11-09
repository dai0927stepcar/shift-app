import io
import calendar
import datetime
import random
import hashlib
import json
import pandas as pd
import streamlit as st
from ortools.sat.python import cp_model
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="シフト自動作成システム", layout="wide")
st.title("📅 シフト自動作成システム")

# =====================================================================
# 🔒 パスワードロック機能（公開向けセキュリティ対応）
# =====================================================================
try:
    PASSWORD = st.secrets["APP_PASSWORD"]
except Exception:
    st.error("⚠️ **【セキュリティエラー】**\n\n"
             "本番環境用のパスワードが設定されていません。\n"
             "不正アクセスを防止するため、システムの起動を停止しました。\n\n"
             "**[管理者の方へ]**\n"
             "`.streamlit/secrets.toml` ファイルを作成し、以下のようにパスワードを設定してください。\n"
             "```toml\nAPP_PASSWORD = \"任意の強固なパスワード\"\n```")
    st.stop()

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
S_HAYA, S_Q, S_OSO, S_IRI, S_AKE, S_YASU, S_YUKU, S_KEN, S_FREE, S_OJT = range(10)

shift_str = {
    S_HAYA: '早', S_Q: 'Q', S_OSO: '遅', S_IRI: '入', 
    S_AKE: '明', S_YASU: '休', S_YUKU: '有給', S_KEN: '研修', S_FREE: 'フリー', S_OJT: 'OJT'
}

# 内部計算のためにすべての表記を同じシフトIDにマッピング（旧表記は互換性維持のため残存）
str_to_shift = {
    '早': S_HAYA, '早①': S_HAYA, '早②': S_HAYA, '早③': S_HAYA,
    'Q': S_Q,
    '遅': S_OSO, '遅①': S_OSO, '遅②': S_OSO, '遅③': S_OSO,
    '入': S_IRI, 'A夜': S_IRI, 'B夜': S_IRI,
    '明': S_AKE, 'A明': S_AKE, 'B明': S_AKE,
    '休': S_YASU, '希望休': S_YASU,
    '有給': S_YUKU, '研修': S_KEN, 'フリー': S_FREE, 'OJT': S_OJT
}

# 新規作成・プルダウン用（旧表記なし）
shift_options = ["", "早①", "早②", "早③", "Q", "遅①", "遅②", "遅③", "A夜", "B夜", "A明", "B明", "希望休", "休", "有給", "研修", "OJT", "フリー"]
allowed_shifts = set(shift_options)

limit_options = ["制限なし"] + [str(i) for i in range(0, 32)]
prev_work_options = [str(i) for i in range(0, 6)]
PREV_COL = "前月連勤"
WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]

DEFAULT_STAFF_NAMES = [chr(ord('A') + i) for i in range(16)]
DEFAULT_Q_DAYS = [1, 3, 6, 7, 8, 13, 14, 15, 17, 20, 22, 24, 27, 28, 29]

def get_default_staff_df():
    data = []
    for name in DEFAULT_STAFF_NAMES:
        data.append({"スタッフ名": name, "早の上限": "制限なし", "Qの上限": "制限なし", "遅の上限": "制限なし", "入の上限": "制限なし", "属性1": "", "属性2": ""})
    return pd.DataFrame(data)

def get_default_ng_pairs_df():
    return pd.DataFrame({"スタッフ1": pd.Series(dtype="str"), "スタッフ2": pd.Series(dtype="str")})

def get_default_personal_lock_df():
    return pd.DataFrame({"スタッフ名": pd.Series(dtype="str"), "開始日": pd.Series(dtype="str"), "終了日": pd.Series(dtype="str")})

def get_default_committee_df():
    return pd.DataFrame({"委員会名": pd.Series(dtype="str"), "対象シフト（例: 遅①,A明）": pd.Series(dtype="str")})

def extract_day_num(col_name):
    s = str(col_name).strip()
    if "(" in s: s = s.split("(")[0].strip()
    return int(s) if s.isdigit() else None

def apply_excel_dropdowns(writer, staffs_count, target_days):
    shift_list = '"早①,早②,早③,Q,遅①,遅②,遅③,A夜,B夜,A明,B明,希望休,休,有給,研修,OJT,フリー"'
    
    dv = DataValidation(type="list", formula1=shift_list, allow_blank=True, 
                        showErrorMessage=True, errorStyle="stop", 
                        errorTitle="入力エラー", error="リストから指定の勤務区分を選択してください。旧表記（早・遅など）は禁止されています。")
    
    if '希望休入力' in writer.sheets:
        ws_req = writer.sheets['希望休入力']
        ws_req.add_data_validation(dv)
        dv.add(f"C2:{get_column_letter(2 + target_days)}{1 + staffs_count}")
        
        dv_prev = DataValidation(type="list", formula1='"0,1,2,3,4,5"', allow_blank=True, showErrorMessage=True, errorStyle="stop", errorTitle="入力エラー", error="0〜5の整数を入力してください。")
        ws_req.add_data_validation(dv_prev)
        dv_prev.add(f"B2:B{1 + staffs_count}")
        
    if '完成シフト表' in writer.sheets:
        ws_res = writer.sheets['完成シフト表']
        dv_res = DataValidation(type="list", formula1=shift_list, allow_blank=True, 
                                showErrorMessage=True, errorStyle="stop", 
                                errorTitle="入力エラー", error="リストから指定の勤務区分を選択してください。")
        ws_res.add_data_validation(dv_res)
        dv_res.add(f"B2:{get_column_letter(1 + target_days)}{1 + staffs_count}")
        
    if '診断用シフト表（未確定）' in writer.sheets:
        ws_diag = writer.sheets['診断用シフト表（未確定）']
        dv_diag = DataValidation(type="list", formula1=shift_list, allow_blank=True, showErrorMessage=False)
        ws_diag.add_data_validation(dv_diag)
        dv_diag.add(f"B2:{get_column_letter(1 + target_days)}{1 + staffs_count}")

# セッション状態の初期化
now = datetime.datetime.now()
if "editor_key" not in st.session_state: st.session_state.editor_key = 0
if "setting_year" not in st.session_state: st.session_state.setting_year = now.year
if "setting_month" not in st.session_state: st.session_state.setting_month = now.month

# 日数は常にカレンダーから自動計算
_, auto_d = calendar.monthrange(st.session_state.setting_year, st.session_state.setting_month)
st.session_state.setting_days = auto_d

if "setting_yasu" not in st.session_state: st.session_state.setting_yasu = 9
if "setting_q_days" not in st.session_state:
    st.session_state.setting_q_days = [d for d in DEFAULT_Q_DAYS if d <= auto_d]
if "exec_mode" not in st.session_state: st.session_state.exec_mode = "通常作成モード（高速・自動3段階緩和）"
if "lock_until_day" not in st.session_state: st.session_state.lock_until_day = 0
if "base_schedule_df" not in st.session_state: st.session_state.base_schedule_df = None
if "staff_names" not in st.session_state: st.session_state.staff_names = DEFAULT_STAFF_NAMES.copy()

# キャッシュ保持変数群
if "assigned_result" not in st.session_state: st.session_state.assigned_result = None
if "assigned_comm" not in st.session_state: st.session_state.assigned_comm = None
if "assigned_changes" not in st.session_state: st.session_state.assigned_changes = None
if "assigned_diag" not in st.session_state: st.session_state.assigned_diag = None
if "last_calc_hash" not in st.session_state: st.session_state.last_calc_hash = ""
if "last_executed_mode" not in st.session_state: st.session_state.last_executed_mode = ""

if "staff_df" not in st.session_state: st.session_state.staff_df = get_default_staff_df()
if "ng_pairs_df" not in st.session_state: st.session_state.ng_pairs_df = get_default_ng_pairs_df()
if "personal_lock_df" not in st.session_state: st.session_state.personal_lock_df = get_default_personal_lock_df()
if "committee_df" not in st.session_state: st.session_state.committee_df = get_default_committee_df()
if "req_df" not in st.session_state:
    init_cols = [PREV_COL] + [str(i) for i in range(1, auto_d + 1)]
    init_req = pd.DataFrame("", index=st.session_state.staff_names, columns=init_cols)
    init_req[PREV_COL] = "0"
    st.session_state.req_df = init_req

ek = st.session_state.editor_key

# =====================================================================
# 設定変更の検知用ハッシュ生成（base_schedule_dfを含む）
# =====================================================================
def get_state_hash():
    state = {
        "year": st.session_state.setting_year,
        "month": st.session_state.setting_month,
        "days": st.session_state.setting_days,
        "yasu": st.session_state.setting_yasu,
        "q_days": sorted(st.session_state.setting_q_days),
        "mode": st.session_state.exec_mode,
        "lock": st.session_state.lock_until_day,
        "staff": st.session_state.staff_df.fillna("").astype(str).to_dict(orient="records"),
        "ng": st.session_state.ng_pairs_df.fillna("").astype(str).to_dict(orient="records"),
        "plock": st.session_state.personal_lock_df.fillna("").astype(str).to_dict(orient="records"),
        "comm": st.session_state.committee_df.fillna("").astype(str).to_dict(orient="records"),
        "req": st.session_state.req_df.fillna("").astype(str).to_dict(orient="index"),
        "base_sched": st.session_state.base_schedule_df.fillna("").astype(str).to_dict(orient="index") if st.session_state.base_schedule_df is not None else None
    }
    s = json.dumps(state, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(s.encode('utf-8')).hexdigest()

# =====================================================================
# 職員名の同期機能（名前が変わっても別人にコピーされない安全な処理）
# =====================================================================
def sync_staff_names(new_names):
    if len(new_names) != len(set(new_names)):
        return False, "❌ スタッフ名が重複しています。別々の名前に変更してください。"
    
    if new_names != st.session_state.staff_names:
        old_req = st.session_state.req_df
        new_req = pd.DataFrame("", index=new_names, columns=old_req.columns)
        new_req[PREV_COL] = "0"
        for name in new_names:
            if name in old_req.index:
                new_req.loc[name] = old_req.loc[name]
        st.session_state.req_df = new_req
        st.session_state.staff_names = new_names
        
        # 名前が変わった場合は基準シフトの互換性も失われるためリセットする
        if st.session_state.base_schedule_df is not None:
            if set(new_names) != set(st.session_state.base_schedule_df.index):
                st.session_state.base_schedule_df = None
                
        st.session_state.editor_key += 1
        st.session_state.assigned_result = None
        return True, ""
    return False, ""

# =====================================================================
# Excel検証と一括読み込み（トランザクション処理）
# =====================================================================
def verify_and_load_excel(xls_file, is_base_schedule=False):
    errors = []
    try:
        xls = pd.ExcelFile(xls_file)
    except Exception:
        return ["Excelファイルの読み込みに失敗しました。ファイルが破損している可能性があります。"], None
        
    temp_state = {}
    
    if "基本設定" not in xls.sheet_names:
        errors.append("「基本設定」シートが存在しません。")
    else:
        df_set = pd.read_excel(xls, "基本設定")
        settings = dict(zip(df_set["設定項目"].astype(str), df_set["設定値"]))
        try:
            y_val = int(float(settings.get("対象年", now.year)))
            m_val = int(float(settings.get("対象月", now.month)))
            d_val = int(float(settings.get("月の日数", 31)))
            yasu_val = int(float(settings.get("休の日数", 9)))
            q_raw = str(settings.get("Qを配置しない日", ""))
            q_days_val = [] if not q_raw or q_raw.lower()=="nan" or q_raw=="なし" else [int(x.strip().split(".")[0]) for x in q_raw.split(",") if x.strip().split(".")[0].isdigit()]
            
            if not (2000 <= y_val <= 2100): errors.append("対象年が有効な範囲外です。")
            if not (1 <= m_val <= 12): errors.append("対象月が1〜12の範囲外です。")
            
            if not errors:
                _, cal_days = calendar.monthrange(y_val, m_val)
                if d_val != cal_days:
                    errors.append(f"設定された対象月（{y_val}年{m_val}月）の日数（{d_val}日）が実際のカレンダー（{cal_days}日）と一致しません。")
                
            temp_state['year'] = y_val
            temp_state['month'] = m_val
            temp_state['days'] = cal_days if not errors else 31
            temp_state['yasu'] = yasu_val
            temp_state['q_days'] = q_days_val
            
        except ValueError:
            errors.append("基本設定シートの数値の読み取りに失敗しました。")

    if "スタッフ設定" not in xls.sheet_names:
        errors.append("「スタッフ設定」シートが存在しません。")
    else:
        df_staff = pd.read_excel(xls, "スタッフ設定").fillna("制限なし").astype(str)
        if "スタッフ名" not in df_staff.columns:
            errors.append("スタッフ設定シートに「スタッフ名」列がありません。")
        else:
            names = [n.strip() for n in df_staff["スタッフ名"] if n.strip()]
            if not names:
                errors.append("スタッフ名が1件も登録されていません。")
            if len(names) != len(set(names)):
                errors.append("スタッフ名に重複があります。同名の場合は番号等を付けて一意にしてください。")
            
            if "属性1" not in df_staff.columns: df_staff["属性1"] = ""
            if "属性2" not in df_staff.columns: df_staff["属性2"] = ""
            df_staff["属性1"] = df_staff["属性1"].replace("制限なし", "")
            df_staff["属性2"] = df_staff["属性2"].replace("制限なし", "")
            temp_state['staff_df'] = df_staff
            temp_state['staff_names'] = names

    if "希望休入力" not in xls.sheet_names:
        errors.append("「希望休入力」シートが存在しません（前月連勤の読み取りに必要です）。")

    target_sheet = "完成シフト表" if is_base_schedule else "希望休入力"
    if is_base_schedule and target_sheet not in xls.sheet_names:
        errors.append(f"「{target_sheet}」シートが存在しません。")

    if not errors:
        df_req_sheet = pd.read_excel(xls, "希望休入力", index_col=0).fillna("").astype(str)
        df_target_sheet = pd.read_excel(xls, target_sheet, index_col=0).fillna("").astype(str)
        
        req_names = [str(n).strip() for n in df_req_sheet.index if str(n).strip()]
        target_names = [str(n).strip() for n in df_target_sheet.index if str(n).strip()]
        
        if set(req_names) != set(temp_state['staff_names']):
            errors.append(f"「希望休入力」の職員名と「スタッフ設定」の職員名が一致しません。")
        if is_base_schedule and set(target_names) != set(temp_state['staff_names']):
            errors.append(f"「完成シフト表」の職員名と「スタッフ設定」の職員名が一致しません。")
            
        norm_cols = [PREV_COL] + [str(i) for i in range(1, temp_state['days'] + 1)]
        
        # 必須列（前月連勤）は希望休シートから探す
        req_prev_col = None
        for col in df_req_sheet.columns:
            if str(col).strip() == PREV_COL: req_prev_col = col
        if not req_prev_col:
            errors.append(f"希望休入力シートに必須列「{PREV_COL}」が見つかりません。")
            
        extracted_t_cols = {}
        for col in df_target_sheet.columns:
            if str(col).strip() == PREV_COL: continue
            d_num = extract_day_num(col)
            if d_num and 1 <= d_num <= temp_state['days']:
                if str(d_num) in extracted_t_cols:
                    errors.append(f"「{target_sheet}」で日付列「{d_num}」が重複して存在しています。")
                extracted_t_cols[str(d_num)] = col
                
        for req_col in norm_cols[1:]:
            if req_col not in extracted_t_cols:
                errors.append(f"「{target_sheet}」に日付列「{req_col}」が見つかりません。")
                
        if not errors:
            normalized_req = pd.DataFrame("", index=temp_state['staff_names'], columns=norm_cols)
            normalized_base = pd.DataFrame("", index=temp_state['staff_names'], columns=norm_cols) if is_base_schedule else None
            
            for s_name in temp_state['staff_names']:
                v_prev = str(df_req_sheet.loc[s_name, req_prev_col]).strip()
                if v_prev.isdigit() and 0 <= int(v_prev) <= 5:
                    normalized_req.loc[s_name, PREV_COL] = str(int(v_prev))
                    if is_base_schedule: normalized_base.loc[s_name, PREV_COL] = str(int(v_prev))
                else:
                    errors.append(f"{s_name}さんの前月連勤の値「{v_prev}」が不正です。0〜5の整数を指定してください。")
                    
                for d in range(1, temp_state['days'] + 1):
                    # req_df 用の値
                    req_val = str(df_req_sheet.loc[s_name, extracted_t_cols[str(d)]] if not is_base_schedule else df_target_sheet.loc[s_name, extracted_t_cols[str(d)]]).strip()
                    if req_val and req_val not in allowed_shifts:
                        errors.append(f"{target_sheet}：{s_name}さんの{d}日の入力値「{req_val}」は不正な勤務区分です（旧表記は使用不可）。")
                    normalized_req.loc[s_name, str(d)] = req_val
                    
                    # base_schedule_df 用の値（欠勤フォロー時）
                    if is_base_schedule:
                        base_val = str(df_target_sheet.loc[s_name, extracted_t_cols[str(d)]]).strip()
                        if base_val and base_val not in allowed_shifts:
                            errors.append(f"完成シフト表：{s_name}さんの{d}日の入力値「{base_val}」は不正な勤務区分です。")
                        normalized_base.loc[s_name, str(d)] = base_val

            temp_state['req_df'] = normalized_req
            if is_base_schedule:
                temp_state['base_schedule_df'] = normalized_base
            else:
                temp_state['base_schedule_df'] = None

    if not errors:
        if "NGペア設定" in xls.sheet_names:
            df_ng = pd.read_excel(xls, "NGペア設定").fillna("").astype(str)
            if "スタッフ1" in df_ng.columns and "スタッフ2" in df_ng.columns: temp_state['ng_pairs_df'] = df_ng[["スタッフ1", "スタッフ2"]]
        if "個別期間ロック設定" in xls.sheet_names:
            df_pl = pd.read_excel(xls, "個別期間ロック設定").fillna("").astype(str)
            if "スタッフ名" in df_pl.columns and "開始日" in df_pl.columns and "終了日" in df_pl.columns:
                temp_state['personal_lock_df'] = df_pl[["スタッフ名", "開始日", "終了日"]]
        if "委員会設定" in xls.sheet_names:
            df_cm = pd.read_excel(xls, "委員会設定").fillna("").astype(str)
            if "委員会名" in df_cm.columns and "対象シフト（例: 遅①,A明）" in df_cm.columns:
                temp_state['committee_df'] = df_cm[["委員会名", "対象シフト（例: 遅①,A明）"]]

    return errors, temp_state if not errors else None

# =====================================================================
# 手動指定の矛盾事前チェック（計算前）
# =====================================================================
def validate_manual_inputs(req_df, staff_list, target_days):
    fatal_input_errors = []
    
    for d in range(1, target_days + 1):
        day_reqs = [str(req_df.loc[s_name, str(d)]).strip() for s_name in staff_list if str(d) in req_df.columns]
        
        for role in ['早①','早②','早③', '遅①','遅②','遅③', 'A夜','B夜', 'A明','B明']:
            if day_reqs.count(role) > 1:
                fatal_input_errors.append(f"{d}日で「{role}」の手動指定が重複しています。")

    for s_name in staff_list:
        for d in range(1, target_days):
            req_curr = str(req_df.loc[s_name, str(d)]).strip() if str(d) in req_df.columns else ""
            req_nxt = str(req_df.loc[s_name, str(d+1)]).strip() if str(d+1) in req_df.columns else ""
            
            if req_curr in ['A夜', 'B夜'] and req_nxt in ['A明', 'B明']:
                exp = req_curr.replace('夜', '明')
                if req_nxt != exp:
                    fatal_input_errors.append(f"{s_name}さんの{d}日指定（{req_curr}）と翌日指定（{req_nxt}）が矛盾しています。")

    return fatal_input_errors

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
    
    # 存在しない日付のデータを切り落とすため req_df を再構築
    old_req = st.session_state.req_df
    new_cols = [PREV_COL] + [str(i) for i in range(1, max_d + 1)]
    new_req = pd.DataFrame("", index=st.session_state.staff_names, columns=new_cols)
    for s_name in st.session_state.staff_names:
        if s_name in old_req.index:
            for c in new_cols:
                if c in old_req.columns:
                    new_req.loc[s_name, c] = old_req.loc[s_name, c]
    st.session_state.req_df = new_req
    
    st.session_state.setting_q_days = [d for d in st.session_state.setting_q_days if d <= max_d]
    st.session_state.base_schedule_df = None
    st.session_state.editor_key += 1
    st.session_state.assigned_result = None
    st.rerun()

target_year = st.session_state.setting_year
target_month = st.session_state.setting_month
TARGET_DAYS = st.session_state.setting_days
st.sidebar.info(f"月の日数: {TARGET_DAYS}日 (カレンダー自動計算)")

day_labels = {}
for d in range(1, TARGET_DAYS + 1):
    wd = WEEKDAYS_JP[datetime.date(target_year, target_month, d).weekday()]
    day_labels[d] = f"{d}({wd})"

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
    st.sidebar.info("💡 **欠勤フォロー手順**\n1. 「完成シフト表」を読み込む。\n2. 下の表で欠勤者の該当日のシフトを書き換える。\n3. すでに終わった日（過去日）を下で指定して作成ボタンを押す。")
    LOCK_UNTIL_DAY = int(st.sidebar.number_input("🔒 何日目までを過去日として完全固定するか（0＝固定なし）", min_value=0, max_value=TARGET_DAYS - 1, value=min(int(st.session_state.lock_until_day), TARGET_DAYS - 1)))
    st.session_state.lock_until_day = LOCK_UNTIL_DAY
else:
    LOCK_UNTIL_DAY = 0

st.sidebar.divider()
st.sidebar.subheader("📂 保存データの読み込み")

uploaded_file = st.sidebar.file_uploader("Excelファイルを選択", type=["xlsx"])
if uploaded_file is not None:
    if st.sidebar.button("📥 入力途中のデータ（希望休）を復元", use_container_width=True):
        errs, t_state = verify_and_load_excel(uploaded_file, is_base_schedule=False)
        if errs:
            st.sidebar.error("❌ 読み込みエラー\n" + "\n".join(errs))
        else:
            st.session_state.setting_year = t_state['year']
            st.session_state.setting_month = t_state['month']
            st.session_state.setting_days = t_state['days']
            st.session_state.setting_yasu = t_state['yasu']
            st.session_state.setting_q_days = t_state['q_days']
            st.session_state.staff_df = t_state['staff_df']
            st.session_state.staff_names = t_state['staff_names']
            st.session_state.req_df = t_state['req_df']
            st.session_state.base_schedule_df = None
            if 'ng_pairs_df' in t_state: st.session_state.ng_pairs_df = t_state['ng_pairs_df']
            if 'personal_lock_df' in t_state: st.session_state.personal_lock_df = t_state['personal_lock_df']
            if 'committee_df' in t_state: st.session_state.committee_df = t_state['committee_df']
            
            st.session_state.exec_mode = "通常作成モード（高速・自動3段階緩和）"
            st.session_state.editor_key += 1
            st.session_state.assigned_result = None
            st.sidebar.success("✅ 希望休・設定データを復元しました！")
            st.rerun()

    if st.sidebar.button("🛠️ 完成シフト表を読み込む（欠勤フォロー用）", type="primary", use_container_width=True):
        errs, t_state = verify_and_load_excel(uploaded_file, is_base_schedule=True)
        if errs:
            st.sidebar.error("❌ 完成シフト表の読み込みエラー\n" + "\n".join(errs))
        else:
            st.session_state.setting_year = t_state['year']
            st.session_state.setting_month = t_state['month']
            st.session_state.setting_days = t_state['days']
            st.session_state.setting_yasu = t_state['yasu']
            st.session_state.setting_q_days = t_state['q_days']
            st.session_state.staff_df = t_state['staff_df']
            st.session_state.staff_names = t_state['staff_names']
            st.session_state.req_df = t_state['req_df']
            st.session_state.base_schedule_df = t_state['base_schedule_df']
            if 'ng_pairs_df' in t_state: st.session_state.ng_pairs_df = t_state['ng_pairs_df']
            if 'personal_lock_df' in t_state: st.session_state.personal_lock_df = t_state['personal_lock_df']
            if 'committee_df' in t_state: st.session_state.committee_df = t_state['committee_df']
            
            st.session_state.exec_mode = "🛠️ 欠勤フォロー（最小変更）モード"
            st.session_state.editor_key += 1
            st.session_state.assigned_result = None
            st.sidebar.success("✅ 完成シフト表を読み込み、欠勤フォローモードに切り替えました！")
            st.rerun()

# =====================================================================
# 1. スタッフ名・個人条件・NGペア等の設定
# =====================================================================
st.subheader("1. スタッフ名・個人条件の設定")
with st.expander("👥 スタッフ名（個人名への書き換え）・シフト上限・属性（委員会）を編集する", expanded=False):
    with st.popover("📝 スタッフ名をまとめて一括変更する（コピペ入力）"):
        bulk_names_text = st.text_area("スタッフ名リストを入力（例：佐藤、鈴木、高橋...）", height=150)
        if st.button("名前を一括反映する", type="primary"):
            raw_list = [n.strip() for n in bulk_names_text.replace("、", "\n").replace(",", "\n").replace(" ", "\n").replace(" ", "\n").splitlines() if n.strip()]
            if raw_list:
                is_ok, err_msg = sync_staff_names(raw_list)
                if not is_ok:
                    st.error(err_msg)
                else:
                    new_staff_df = st.session_state.staff_df.copy()
                    for idx, new_name in enumerate(raw_list):
                        if idx < len(new_staff_df): new_staff_df.loc[idx, "スタッフ名"] = new_name
                        else: new_staff_df = pd.concat([new_staff_df, pd.DataFrame([{"スタッフ名": new_name, "早の上限": "制限なし", "Qの上限": "制限なし", "遅の上限": "制限なし", "入の上限": "制限なし", "属性1": "", "属性2": ""}])], ignore_index=True)
                    st.session_state.staff_df = new_staff_df
                    st.rerun()

    staff_col_config = {
        "スタッフ名": st.column_config.TextColumn("スタッフ名", required=True, width="medium"),
        "早の上限": st.column_config.SelectboxColumn("早の上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "Qの上限": st.column_config.SelectboxColumn("Qの上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "遅の上限": st.column_config.SelectboxColumn("遅の上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "入の上限": st.column_config.SelectboxColumn("入の上限 (0=禁止)", options=limit_options, default="制限なし", required=True),
        "属性1": st.column_config.TextColumn("属性1 (委員会等)", default=""),
        "属性2": st.column_config.TextColumn("属性2 (委員会等)", default=""),
    }
    edited_staff_df = st.data_editor(st.session_state.staff_df, column_config=staff_col_config, num_rows="dynamic", use_container_width=True, height=min(420, 65 + len(st.session_state.staff_df) * 35), key=f"staff_editor_{ek}")

# data_editor 内で名前が変更・追加・削除された場合の安全な同期処理
current_names_from_editor = [str(n).strip() for n in edited_staff_df["スタッフ名"] if str(n).strip()]
if current_names_from_editor != st.session_state.staff_names:
    is_ok, err_msg = sync_staff_names(current_names_from_editor)
    if not is_ok:
        st.error(err_msg)
    else:
        st.session_state.staff_df = edited_staff_df
        st.rerun()

staff_names = st.session_state.staff_names
staffs = len(staff_names)

with st.expander("🚫 夜勤（入）の同日NGペア設定", expanded=False):
    ng_col_config = {
        "スタッフ1": st.column_config.SelectboxColumn("スタッフ1", options=staff_names, required=True, width="medium"),
        "スタッフ2": st.column_config.SelectboxColumn("スタッフ2", options=staff_names, required=True, width="medium"),
    }
    edited_ng_pairs_df = st.data_editor(st.session_state.ng_pairs_df, column_config=ng_col_config, num_rows="dynamic", use_container_width=True, key=f"ng_editor_{ek}")
st.session_state.ng_pairs_df = edited_ng_pairs_df

with st.expander("📋 委員会・グループ設定（同一日に特定シフトで集める）", expanded=False):
    st.caption("・対象シフトはカンマ区切りで入力します（例：遅①, A明）。\n※番号付きで入力しても、判定時は勤務区分(早番・遅番など全体)が対象になります。\n・対象者の過半数が月内の1日でこのシフトに配置されるようAIが探索します。")
    comm_col_config = {
        "委員会名": st.column_config.TextColumn("委員会名", required=True),
        "対象シフト（例: 遅①,A明）": st.column_config.TextColumn("対象シフト", required=True),
    }
    edited_committee_df = st.data_editor(st.session_state.committee_df, column_config=comm_col_config, num_rows="dynamic", use_container_width=True, key=f"comm_editor_{ek}")
st.session_state.committee_df = edited_committee_df

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
for s_name in staff_names:
    if s_name in prev_req.index:
        for c in all_req_cols:
            if c in prev_req.columns:
                val = str(prev_req.loc[s_name, c]).strip()
                current_req_df.loc[s_name, c] = val if (c != PREV_COL or val in prev_work_options) else "0"

column_configuration = {PREV_COL: st.column_config.SelectboxColumn("前月連勤", options=prev_work_options, default="0", width="small", required=True)}
for i in range(1, TARGET_DAYS + 1):
    col_label = f"🔒{day_labels[i]}" if ("欠勤フォロー" in mode and i <= LOCK_UNTIL_DAY) else day_labels[i]
    column_configuration[str(i)] = st.column_config.SelectboxColumn(label=col_label, options=shift_options, default="", width="small")

edited_req_df = st.data_editor(current_req_df, column_config=column_configuration, use_container_width=True, height=min(620, 65 + staffs * 35), key=f"req_editor_{ek}")
st.session_state.req_df = edited_req_df

# ★欠勤フォローモードで、基準シフトが存在しない場合はエラーを表示して計算をさせない
if "欠勤フォロー" in mode and st.session_state.base_schedule_df is None:
    st.error("❌ 欠勤フォローモードを実行するための「基準となる完成シフト表」が設定されていません。左のメニューから現行の完成シフト表Excelを読み込んでください。")

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
    edited_committee_df.to_excel(writer, sheet_name='委員会設定', index=False)
    edited_personal_lock_df.to_excel(writer, sheet_name='個別期間ロック設定', index=False)
    save_req_df.to_excel(writer, sheet_name='希望休入力')
    apply_excel_dropdowns(writer, staffs, TARGET_DAYS)

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
            st.session_state.assigned_result = None
            st.rerun()

st.sidebar.download_button("💾 現在の入力進捗を保存", data=save_excel_bytes, file_name=f"shift_progress_{target_year}_{target_month:02d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
st.sidebar.divider()
if st.sidebar.button("🔓 ログアウト"):
    st.session_state.authenticated = False
    st.rerun()

# =====================================================================
# 画面データ読み取りと状態の検証
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

def get_committees():
    committees = []
    for _, row in edited_committee_df.iterrows():
        c_name = str(row.get("委員会名", "")).strip()
        shifts_str = str(row.get("対象シフト（例: 遅①,A明）", "")).strip()
        if c_name and shifts_str:
            members = []
            for s_idx in range(staffs):
                attr1 = str(edited_staff_df.iloc[s_idx].get("属性1", "")).strip()
                attr2 = str(edited_staff_df.iloc[s_idx].get("属性2", "")).strip()
                if c_name == attr1 or c_name == attr2:
                    members.append(s_idx)
            
            target_shifts = []
            parsed_shifts = shifts_str.replace("、", ",").replace("・", ",").replace(" ", ",").replace(" ", ",")
            for s_name_part in parsed_shifts.split(","):
                sn = s_name_part.strip()
                if sn in str_to_shift:
                    target_shifts.append(str_to_shift[sn])
            
            # ★修正1：同じ勤務区分IDの重複登録を除去
            if members and target_shifts:
                unique_shifts = list(dict.fromkeys(target_shifts))
                req_count = len(members) // 2 + 1
                committees.append((c_name, members, unique_shifts, req_count))
    return committees

PERSONAL_LOCKS = get_personal_locks()
COMMITTEES = get_committees()
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
        
        if orig_str in str_to_shift:
            ORIGINAL_SHIFTS[(s_idx, d)] = str_to_shift[orig_str]
            ORIGINAL_STR_SHIFTS[(s_idx, d)] = orig_str

        is_personal_locked = any(s == s_idx and start <= d <= end for s, start, end in PERSONAL_LOCKS)

        if "欠勤フォロー" not in mode:
            if val_str in str_to_shift: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]
        else:
            if val_str in str_to_shift:
                if d <= LOCK_UNTIL_DAY: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]
                elif is_personal_locked: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]
                elif orig_str != "" and val_str != orig_str: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]
                elif val_str in ["希望休", "有給", "研修", "OJT", "フリー"]: SPECIFIC_SHIFTS[(s_idx, d)] = str_to_shift[val_str]

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

def is_explicit_fixed(s_name, d):
    req_val = str(edited_req_df.loc[s_name, str(d)]).strip() if str(d) in edited_req_df.columns else ""
    if not req_val: return False, ""
    
    if "欠勤フォロー" in mode:
        base_val = ""
        if base_df is not None and s_name in base_df.index and str(d) in base_df.columns:
            base_val = str(base_df.loc[s_name, str(d)]).strip()
        
        s_idx = staff_names.index(s_name)
        is_p_locked = any(s == s_idx and start <= d <= end for s, start, end in PERSONAL_LOCKS)
        is_past_locked = (d <= LOCK_UNTIL_DAY)
        
        if is_p_locked or is_past_locked: return True, req_val
        if req_val != base_val: return True, req_val
        if req_val in ["希望休", "休", "有給", "研修", "OJT", "フリー"]: return True, req_val
            
        return False, ""
    else:
        return True, req_val

# =====================================================================
# ★表示用役割名の割り当てと検証（旧表記完全排除版）
# =====================================================================
def assign_role_names(raw_table_data, req_df):
    fatal_errors = []
    
    for s in range(staffs):
        for d in range(1, TARGET_DAYS):
            is_f_iri, r_iri = is_explicit_fixed(staff_names[s], d)
            is_f_ake, r_ake = is_explicit_fixed(staff_names[s], d+1)
            iri_l = r_iri[0] if is_f_iri and r_iri in ['A夜', 'B夜'] else None
            ake_l = r_ake[0] if is_f_ake and r_ake in ['A明', 'B明'] else None
            if iri_l and ake_l and iri_l != ake_l:
                fatal_errors.append(f"{staff_names[s]}さんの{d}日夜勤入り（{r_iri}）と翌日夜勤明け（{r_ake}）の固定指定が矛盾しています。")

    ake_staffs_d1 = [s for s in range(staffs) if raw_table_data[s][0] == '明']
    req_ake_counts = {'A明': [], 'B明': []}
    for s in ake_staffs_d1:
        is_f, req_val = is_explicit_fixed(staff_names[s], 1)
        if is_f and req_val in req_ake_counts:
            req_ake_counts[req_val].append(staff_names[s])
    if len(req_ake_counts['A明']) > 1:
        fatal_errors.append(f"1日の夜勤明けで「A明」の固定指定が重複しています（{', '.join(req_ake_counts['A明'])}）。")
    if len(req_ake_counts['B明']) > 1:
        fatal_errors.append(f"1日の夜勤明けで「B明」の固定指定が重複しています（{', '.join(req_ake_counts['B明'])}）。")

    if fatal_errors:
        return fatal_errors

    for d_idx in range(TARGET_DAYS):
        d = d_idx + 1
        
        haya_staffs = [s for s in range(staffs) if raw_table_data[s][d_idx] == '早']
        oso_staffs  = [s for s in range(staffs) if raw_table_data[s][d_idx] == '遅']
        iri_staffs  = [s for s in range(staffs) if raw_table_data[s][d_idx] == '入']
        ake_staffs  = [s for s in range(staffs) if raw_table_data[s][d_idx] == '明']
        
        def assign_roles(staff_list, available_roles, base_role, check_prev_maintain=False):
            assigned = {}
            avail = available_roles.copy()
            unassigned = []
            req_set = set()
            
            for s in staff_list:
                is_fixed, req_val = is_explicit_fixed(staff_names[s], d)
                if is_fixed and req_val in available_roles:
                    if req_val in req_set:
                        fatal_errors.append(f"{d}日の{base_role}番で「{req_val}」の指定が重複しています。")
                    else:
                        req_set.add(req_val)
                        assigned[s] = req_val
                        avail.remove(req_val)
                else:
                    unassigned.append(s)
                    
            if check_prev_maintain and base_df is not None:
                random.shuffle(unassigned)
                for s in unassigned[:]:
                    orig_s = ORIGINAL_STR_SHIFTS.get((s, d))
                    if orig_s in avail:
                        assigned[s] = orig_s
                        avail.remove(orig_s)
                        unassigned.remove(s)
                        
            random.shuffle(unassigned)
            for s in unassigned:
                if avail:
                    assigned[s] = avail.pop(0)
                else:
                    assigned[s] = f"{base_role}?"

            return assigned

        assigned_haya = assign_roles(haya_staffs, ['早①', '早②', '早③'], '早', check_prev_maintain=True)
        for s, r in assigned_haya.items(): raw_table_data[s][d_idx] = r

        assigned_oso = assign_roles(oso_staffs, ['遅①', '遅②', '遅③'], '遅', check_prev_maintain=True)
        for s, r in assigned_oso.items(): raw_table_data[s][d_idx] = r

        assigned_iri = {}
        avail_iri = ['A夜', 'B夜']
        unassigned_iri = []
        req_iri_set = set()

        for s in iri_staffs:
            is_f_iri, req_val = is_explicit_fixed(staff_names[s], d)
            is_f_ake, nxt_req = False, ""
            if d_idx + 1 < TARGET_DAYS:
                is_f_ake, nxt_req = is_explicit_fixed(staff_names[s], d+1)
                
            r_l = req_val[0] if is_f_iri and req_val in ['A夜', 'B夜'] else None
            n_l = nxt_req[0] if is_f_ake and nxt_req in ['A明', 'B明'] else None
            final_l = r_l or n_l
            req_iri_str = f"{final_l}夜" if final_l else None

            if req_iri_str:
                if req_iri_str in req_iri_set:
                    fatal_errors.append(f"{d}日の夜勤入りで「{req_iri_str}」が重複して指定・推論されました。")
                else:
                    req_iri_set.add(req_iri_str)
                    assigned_iri[s] = req_iri_str
                    if req_iri_str in avail_iri:
                        avail_iri.remove(req_iri_str)
            else:
                unassigned_iri.append(s)

        if "欠勤フォロー" in mode and base_df is not None:
            random.shuffle(unassigned_iri)
            for s in unassigned_iri[:]:
                orig_s = ORIGINAL_STR_SHIFTS.get((s, d))
                if orig_s in avail_iri:
                    assigned_iri[s] = orig_s
                    avail_iri.remove(orig_s)
                    unassigned_iri.remove(s)

        random.shuffle(unassigned_iri)
        for s in unassigned_iri:
            if avail_iri:
                assigned_iri[s] = avail_iri.pop(0)
            else:
                assigned_iri[s] = "入?"

        for s, r in assigned_iri.items(): raw_table_data[s][d_idx] = r

        assigned_ake = {}
        avail_ake = ['A明', 'B明']
        unassigned_ake = []

        for s in ake_staffs:
            if d_idx == 0:
                is_f, req_val = is_explicit_fixed(staff_names[s], d)
                if is_f and req_val in avail_ake:
                    assigned_ake[s] = req_val
                    avail_ake.remove(req_val)
                else:
                    unassigned_ake.append(s)
            else:
                # ★修正4：夜勤明けのアルファベットは、必ず「前日の実際の勤務区分」だけを見て確定する
                prev_shift = raw_table_data[s][d_idx - 1]
                if prev_shift == 'A夜':
                    assigned_str = 'A明'
                elif prev_shift == 'B夜':
                    assigned_str = 'B明'
                else:
                    assigned_str = None
                
                if assigned_str:
                    assigned_ake[s] = assigned_str
                    if assigned_str in avail_ake:
                        avail_ake.remove(assigned_str)
                else:
                    unassigned_ake.append(s)

        if d_idx == 0 and unassigned_ake:
            if "欠勤フォロー" in mode and base_df is not None:
                random.shuffle(unassigned_ake)
                for s in unassigned_ake[:]:
                    orig_s = ORIGINAL_STR_SHIFTS.get((s, d))
                    if orig_s in avail_ake:
                        assigned_ake[s] = orig_s
                        avail_ake.remove(orig_s)
                        unassigned_ake.remove(s)
                        
            random.shuffle(unassigned_ake)
            for s in unassigned_ake:
                if avail_ake:
                    assigned_ake[s] = avail_ake.pop(0)
                else:
                    assigned_ake[s] = "明?"
        elif d_idx > 0 and unassigned_ake:
            for s in unassigned_ake:
                assigned_ake[s] = "明?"
                
        for s, r in assigned_ake.items(): raw_table_data[s][d_idx] = r

    return fatal_errors

def validate_assigned_roles(raw_table_data, original_raw_table):
    fatal_errors = []
    rule_warnings = []
    
    for d_idx in range(TARGET_DAYS):
        d = d_idx + 1
        orig_shifts = [original_raw_table[s][d_idx] for s in range(staffs)]
        assigned_shifts = [raw_table_data[s][d_idx] for s in range(staffs)]
        
        h_cnt = orig_shifts.count('早')
        o_cnt = orig_shifts.count('遅')
        i_cnt = orig_shifts.count('入')
        a_cnt = orig_shifts.count('明')
        
        if h_cnt != 3: rule_warnings.append(f"{d}日の早番が{h_cnt}名です（通常3名）。")
        if o_cnt != 3: rule_warnings.append(f"{d}日の遅番が{o_cnt}名です（通常3名）。")
        if i_cnt != 2: rule_warnings.append(f"{d}日の夜勤入りが{i_cnt}名です（通常2名）。")
        if a_cnt != 2: rule_warnings.append(f"{d}日の夜勤明けが{a_cnt}名です（通常2名）。")
        
        for role in ['早①','早②','早③']:
            if assigned_shifts.count(role) > 1: fatal_errors.append(f"{d}日で「{role}」の割り当てが重複しています。")
        for role in ['遅①','遅②','遅③']:
            if assigned_shifts.count(role) > 1: fatal_errors.append(f"{d}日で「{role}」の割り当てが重複しています。")
        for role in ['A夜','B夜']:
            if assigned_shifts.count(role) > 1: fatal_errors.append(f"{d}日で「{role}」の割り当てが重複しています。")
        for role in ['A明','B明']:
            if assigned_shifts.count(role) > 1: fatal_errors.append(f"{d}日で「{role}」の割り当てが重複しています。")
            
        if any('?' in x for x in assigned_shifts):
            for s, x in enumerate(assigned_shifts):
                if '?' in x:
                    rule_warnings.append(f"{d}日の{staff_names[s]}さんの「{x}」が未確定です。人数制約緩和や手動指定の矛盾が原因の可能性があります。")

    for s in range(staffs):
        for d in range(TARGET_DAYS - 1):
            curr_a = raw_table_data[s][d]
            nxt_a = raw_table_data[s][d+1]
            if curr_a in ['A夜', 'B夜']:
                exp = curr_a.replace('夜', '明')
                if nxt_a != exp:
                    rule_warnings.append(f"{d+1}日の{staff_names[s]}さんが前日{curr_a}ですが、翌日が{nxt_a}になっています。")
        for d in range(1, TARGET_DAYS):
            curr_a = raw_table_data[s][d]
            prv_a = raw_table_data[s][d-1]
            if curr_a in ['A明', 'B明']:
                exp = curr_a.replace('明', '夜')
                if prv_a != exp:
                    rule_warnings.append(f"{d+1}日の{staff_names[s]}さんが{curr_a}ですが、前日が{prv_a}になっています。")

    for s in range(staffs):
        for d in range(TARGET_DAYS):
            is_fixed, req_val = is_explicit_fixed(staff_names[s], d+1)
            if is_fixed and req_val in ['早①', '早②', '早③', '遅①', '遅②', '遅③', 'A夜', 'B夜', 'A明', 'B明']:
                actual = raw_table_data[s][d]
                if actual != req_val:
                    fatal_errors.append(f"{d+1}日の{staff_names[s]}さんの手動指定「{req_val}」が無視され「{actual}」に変更されています。")

    for s in range(staffs):
        for d in range(TARGET_DAYS):
            assigned_val = raw_table_data[s][d].replace('?', '')
            orig_val = original_raw_table[s][d]
            if str_to_shift.get(assigned_val) != str_to_shift.get(orig_val):
                fatal_errors.append(f"{d+1}日の{staff_names[s]}さんについて、内部計算結果({orig_val})と表示用割り当て({assigned_val})が一致しません。")
                    
    return fatal_errors, rule_warnings

def format_result_df(raw_table_data):
    for s_idx, d_idx in KIBOU_YASU_CELLS:
        if 0 <= s_idx < len(raw_table_data) and 0 <= d_idx < TARGET_DAYS:
            if raw_table_data[s_idx][d_idx] == '休': raw_table_data[s_idx][d_idx] = '希望休'
            
    cols = [day_labels[i] for i in range(1, TARGET_DAYS + 1)]
    df = pd.DataFrame(raw_table_data, index=staff_names, columns=cols)
    
    df["【早】"] = df[cols].isin(["早①", "早②", "早③", "早?"]).sum(axis=1)
    df["【Q】"] = (df[cols] == "Q").sum(axis=1)
    df["【遅】"] = df[cols].isin(["遅①", "遅②", "遅③", "遅?"]).sum(axis=1)
    df["【入】"] = df[cols].isin(["A夜", "B夜", "入?"]).sum(axis=1)
    df["【明】"] = df[cols].isin(["A明", "B明", "明?"]).sum(axis=1)
    df["【休計】"] = df[cols].isin(["休", "希望休"]).sum(axis=1)
    df["【有給】"] = (df[cols] == "有給").sum(axis=1)
    if df[cols].isin(["研修"]).any().any(): df["【研修】"] = (df[cols] == "研修").sum(axis=1)
    if df[cols].isin(["OJT"]).any().any(): df["【OJT】"] = (df[cols] == "OJT").sum(axis=1)
    if df[cols].isin(["フリー"]).any().any(): df["【フリー】"] = (df[cols] == "フリー").sum(axis=1)
    return df

def style_shift_df(df):
    def color_cell(val):
        if val in ["休", "希望休"]: return "background-color: #ffcccc; color: #cc0000; font-weight: bold;"
        elif val == "有給": return "background-color: #cce5ff; color: #004085; font-weight: bold;"
        elif "?" in str(val): return "background-color: #fff3cd; color: #856404; font-weight: bold;"
        return ""
    try: return df.style.map(color_cell)
    except AttributeError: return df.style.applymap(color_cell)

def check_shortage():
    if staffs < 10:
        st.error(f"❌ 現在のスタッフ数（{staffs}人）では、1日に必要な勤務枠を満たせません。")
        return False
    
    if "欠勤フォロー" in mode and st.session_state.base_schedule_df is None:
        st.error("❌ 欠勤フォローモードを実行するための「基準となる完成シフト表」が設定されていません。左のメニューから現行の完成シフト表Excelを読み込んでください。")
        return False

    total_slots = staffs * TARGET_DAYS
    work_slots = (3 + 3 + 2 + 2) * TARGET_DAYS + (TARGET_DAYS - len(Q_SHIFT_DAYS))
    specific_other_slots = sum(1 for (s, d), t in SPECIFIC_SHIFTS.items() if s < staffs and t not in [S_YASU, S_HAYA, S_OSO, S_IRI, S_AKE, S_Q])
    shortage = (staffs * TARGET_YASU) - (total_slots - work_slots - specific_other_slots)
    
    if shortage > 0:
        if "通常作成" in mode:
            st.error(f"❌ 【枠不足】全員に「休」を{TARGET_YASU}回配置することが物理的に不可能です。（{shortage} 枠不足しています）")
            return False
        else:
            st.warning(f"⚠️ 【枠不足】全員に「休」を{TARGET_YASU}回配置することが物理的に不可能です。（{shortage} 枠不足しています）\n※不足分は休日出勤として自動調整されます。")
            
    if "欠勤フォロー" in mode: return True
    for s in range(staffs):
        if sum(1 for (s_idx, d_idx), t_idx in SPECIFIC_SHIFTS.items() if s_idx == s and t_idx == S_YASU) > TARGET_YASU:
            if "通常作成" in mode:
                st.error(f"❌ 【希望休超過】{staff_names[s]}さんに休が {TARGET_YASU} 回を超えて指定されています。")
                return False
            else:
                st.warning(f"⚠️ 【希望休超過】{staff_names[s]}さんに休が {TARGET_YASU} 回を超えて指定されています。")
    return True

# =====================================================================
# 計算ロジック1：通常モード
# =====================================================================
def run_normal_mode():
    q_days_idx = [d - 1 for d in Q_SHIFT_DAYS]
    personal_limits = get_personal_limits()
    ng_pairs = get_ng_pairs()
    
    def attempt_fast_solve(max_consec, strict_rank):
        model = cp_model.CpModel()
        B = {(s, d, t): model.NewBoolVar(f'B_{s}_{d}_{t}') for s in range(staffs) for d in range(TARGET_DAYS) for t in range(10)}
        five_consec_penalty_terms = []
        committee_penalties = []

        for s in range(staffs):
            for d in range(TARGET_DAYS):
                model.AddExactlyOne([B[s, d, t] for t in range(10)])
                req_shift = SPECIFIC_SHIFTS.get((s, d+1))
                if req_shift != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
                if req_shift != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
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
            C_vars = [model.NewIntVar(0, 100, f'C_{s}_{d}') for d in range(TARGET_DAYS)]
            for d in range(TARGET_DAYS):
                yasu_flag = model.NewBoolVar(f'yasu_{s}_{d}')
                model.Add(B[s, d, S_YASU] + B[s, d, S_YUKU] == 1).OnlyEnforceIf(yasu_flag)
                model.Add(B[s, d, S_YASU] + B[s, d, S_YUKU] == 0).OnlyEnforceIf(yasu_flag.Not())
                
                prev_val = prev_c if d == 0 else C_vars[d-1]
                w_expr = B[s, d, S_HAYA] + B[s, d, S_OSO] + B[s, d, S_Q] + B[s, d, S_KEN] + B[s, d, S_FREE] + B[s, d, S_OJT] + 2 * B[s, d, S_IRI]
                
                model.Add(C_vars[d] == 0).OnlyEnforceIf(yasu_flag)
                model.Add(C_vars[d] == prev_val + w_expr).OnlyEnforceIf(yasu_flag.Not())
                model.Add(C_vars[d] <= max_consec)
                
                if max_consec == 5:
                    diff = model.NewIntVar(0, 100, f'diff_5c_{s}_{d}')
                    model.AddMaxEquality(diff, [0, C_vars[d] - 4])
                    five_consec_penalty_terms.append(diff)

        for s_idx, s_type, max_c, _ in personal_limits:
            if max_c == 0: model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) == 0)
            else: model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) <= max_c)

        for s1, s2, _ in ng_pairs:
            for d in range(TARGET_DAYS): model.Add(B[s1, d, S_IRI] + B[s2, d, S_IRI] <= 1)

        for c_idx, (c_name, members, t_shifts, req_count) in enumerate(COMMITTEES):
            Z_days = []
            for d in range(TARGET_DAYS):
                Z_d = model.NewBoolVar(f'comm_{c_idx}_{d}')
                model.Add(sum(B[s, d, t] for s in members for t in t_shifts) >= req_count).OnlyEnforceIf(Z_d)
                Z_days.append(Z_d)
            
            met_any = model.NewBoolVar(f'comm_met_{c_idx}')
            model.AddMaxEquality(met_any, Z_days)
            viol = model.NewBoolVar(f'comm_viol_{c_idx}')
            model.Add(met_any == 1).OnlyEnforceIf(viol.Not())
            model.Add(met_any == 0).OnlyEnforceIf(viol)
            
            committee_penalties.append(viol * 50000)

        iri_counts = [model.NewIntVar(0, TARGET_DAYS, f'iri_{s}') for s in range(staffs)]
        for s in range(staffs): model.Add(iri_counts[s] == sum(B[s, d, S_IRI] for d in range(TARGET_DAYS)))
        max_iri = model.NewIntVar(0, TARGET_DAYS, 'max_iri')
        model.AddMaxEquality(max_iri, iri_counts)
        model.Minimize(max_iri * 100 + sum(five_consec_penalty_terms) + sum(committee_penalties))

        solver = cp_model.CpSolver()
        solver.parameters.num_search_workers = 8
        solver.parameters.max_time_in_seconds = 15.0
        return solver.Solve(model), solver, B, max_consec, strict_rank

    used_step = ""
    with st.spinner("高速計算中...（基本4連勤ルール）"): 
        status, solver, B, used_limit, used_strict = attempt_fast_solve(4, True)
        used_step = "基本4連勤ルール"
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        with st.spinner("高速再計算中...（緩和5連勤ルール）"): 
            status, solver, B, used_limit, used_strict = attempt_fast_solve(5, True)
            used_step = "緩和5連勤ルール"
    if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        with st.spinner("高速再計算中...（緩和5連勤・順序フリー）"): 
            status, solver, B, used_limit, used_strict = attempt_fast_solve(5, False)
            used_step = "緩和5連勤・順序フリー"

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        original_raw_table = [[shift_str[next(t for t in range(10) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)]
        copied_table = [row[:] for row in original_raw_table]
        
        assign_fatal_errors = assign_role_names(copied_table, edited_req_df)
        if assign_fatal_errors:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), assign_fatal_errors, [], pd.DataFrame(), "通常作成モード"
            
        fatal_errs, rule_warnings = validate_assigned_roles(copied_table, original_raw_table)
        if fatal_errs:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), fatal_errs, [], pd.DataFrame(), "通常作成モード"
        if rule_warnings:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), rule_warnings, [], pd.DataFrame(), "通常作成モード"

        if used_limit == 4 and used_strict:
            st.success("✅ シフト作成完了！（基本4連勤ルールで成功しました）")
        else:
            st.warning(f"⚠️ 基本の4連勤ルールでは条件が厳しすぎたため、一部ルールを緩和（【{used_step}】を適用）して作成しました。")
            
        res_df = format_result_df(copied_table)
        
        comm_records = []
        rename_dict = {}
        for c_idx, (c_name, members, t_shifts, req_count) in enumerate(COMMITTEES):
            met_day = None
            attendees = []
            for d in range(TARGET_DAYS):
                cnt = 0
                att_names = []
                for s in members:
                    for t in t_shifts:
                        if solver.Value(B[s, d, t]):
                            cnt += 1
                            actual_shift = copied_table[s][d]
                            att_names.append(f"{staff_names[s]}({actual_shift})")
                if cnt >= req_count:
                    met_day = d + 1
                    attendees = att_names
                    break
            
            if met_day is not None:
                day_str = day_labels[met_day]
                comm_records.append({"委員会名": c_name, "実施日": day_str, "参加対象者": "、".join(attendees)})
                if day_str in rename_dict:
                    rename_dict[day_str] = rename_dict[day_str] + f" ★{c_name}"
                else:
                    rename_dict[day_str] = f"{day_str} ★{c_name}"
            else:
                comm_records.append({"委員会名": c_name, "実施日": "⚠️ 条件を満たす日がありません", "参加対象者": "-"})
        comm_res_df = pd.DataFrame(comm_records)

        res_df = res_df.rename(columns=rename_dict)
        return res_df, pd.DataFrame(), comm_res_df, [], [], pd.DataFrame(), "通常作成モード"
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

    def attempt_reschedule(max_consec, strict_rank):
        model = cp_model.CpModel()
        B = {(s, d, t): model.NewBoolVar(f'B_{s}_{d}_{t}') for s in range(staffs) for d in range(TARGET_DAYS) for t in range(10)}
        obj_terms = []
        five_consec_penalty_terms = []

        is_modified = [model.NewBoolVar(f'is_mod_{s}') for s in range(staffs)]

        for s in range(staffs):
            yasu_sum = sum(B[s, d, S_YASU] for d in range(TARGET_DAYS))
            shortage_y = model.NewIntVar(0, TARGET_DAYS, f'shortage_y_{s}')
            model.Add(yasu_sum + shortage_y == TARGET_YASU)
            obj_terms.append(shortage_y * 100000)

            for d in range(TARGET_DAYS):
                model.AddExactlyOne([B[s, d, t] for t in range(10)])
                
                req_shift = SPECIFIC_SHIFTS.get((s, d+1))
                orig_t = ORIGINAL_SHIFTS.get((s, d+1))
                is_p_locked = any(p_s == s and start <= (d+1) <= end for p_s, start, end in PERSONAL_LOCKS)

                if req_shift != S_YUKU and orig_t != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
                if req_shift != S_KEN and orig_t != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
                if req_shift != S_OJT and orig_t != S_OJT: model.Add(B[s, d, S_OJT] == 0)

                if (d + 1) <= LOCK_UNTIL_DAY or is_p_locked:
                    if req_shift is not None:
                        model.Add(B[s, d, req_shift] == 1)
                    elif orig_t is not None:
                        model.Add(B[s, d, orig_t] == 1)
                else:
                    if req_shift == S_YUKU:
                        model.Add(B[s, d, S_YUKU] + B[s, d, S_YASU] == 1)
                        obj_terms.append(B[s, d, S_YASU] * 1)
                    elif req_shift is not None:
                        model.Add(B[s, d, req_shift] == 1)

                    if orig_t is not None:
                        if req_shift == S_YUKU and orig_t == S_YUKU:
                            pass
                        elif req_shift is None:
                            if orig_t in [S_OJT, S_FREE]:
                                obj_terms.append((1 - B[s, d, orig_t]) * 5)
                            elif orig_t in [S_IRI, S_AKE]:
                                obj_terms.append((1 - B[s, d, orig_t]) * 20)
                            elif orig_t == S_YASU:
                                obj_terms.append((1 - B[s, d, S_YASU]) * 30)
                            else:
                                obj_terms.append((1 - B[s, d, orig_t]) * 10)

        for d in range(LOCK_UNTIL_DAY, TARGET_DAYS):
            model.Add(sum(B[s, d, S_HAYA] for s in range(staffs)) == 3)
            model.Add(sum(B[s, d, S_OSO]  for s in range(staffs)) == 3)
            model.Add(sum(B[s, d, S_IRI]  for s in range(staffs)) == 2)
            model.Add(sum(B[s, d, S_AKE]  for s in range(staffs)) == 2)
            if d in q_days_idx: model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 0)
            else: model.Add(sum(B[s, d, S_Q] for s in range(staffs)) == 1)

        for s in range(staffs):
            for d in range(start_check_d, TARGET_DAYS - 1):
                model.AddImplication(B[s, d, S_IRI], B[s, d+1, S_AKE])
                if d + 1 >= LOCK_UNTIL_DAY: model.AddImplication(B[s, d+1, S_AKE], B[s, d, S_IRI])
                model.Add(B[s, d+1, S_YASU] + B[s, d+1, S_YUKU] >= B[s, d, S_AKE])
                if strict_rank:
                    model.AddImplication(B[s, d, S_Q], B[s, d+1, S_HAYA].Not())
                    model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_HAYA].Not())
                    model.AddImplication(B[s, d, S_OSO], B[s, d+1, S_Q].Not())

        for s in range(staffs):
            prev_c = PREV_CONSECUTIVE.get(s, 0)
            C_vars = [model.NewIntVar(0, 100, f'C_{s}_{d}') for d in range(TARGET_DAYS)]
            for d in range(TARGET_DAYS):
                yasu_flag = model.NewBoolVar(f'yasu_{s}_{d}')
                model.Add(B[s, d, S_YASU] + B[s, d, S_YUKU] == 1).OnlyEnforceIf(yasu_flag)
                model.Add(B[s, d, S_YASU] + B[s, d, S_YUKU] == 0).OnlyEnforceIf(yasu_flag.Not())
                
                prev_val = prev_c if d == 0 else C_vars[d-1]
                w_expr = B[s, d, S_HAYA] + B[s, d, S_OSO] + B[s, d, S_Q] + B[s, d, S_KEN] + B[s, d, S_FREE] + B[s, d, S_OJT] + 2 * B[s, d, S_IRI]
                
                model.Add(C_vars[d] == 0).OnlyEnforceIf(yasu_flag)
                model.Add(C_vars[d] == prev_val + w_expr).OnlyEnforceIf(yasu_flag.Not())
                
                if d >= LOCK_UNTIL_DAY:
                    model.Add(C_vars[d] <= max_consec)
                
                diff = model.NewIntVar(0, 100, f'diff_5c_{s}_{d}')
                model.AddMaxEquality(diff, [0, C_vars[d] - 4])
                five_consec_penalty_terms.append(diff)
                
        for term in five_consec_penalty_terms:
            obj_terms.append(term * 50)

        for s_idx, s_type, max_c, _ in personal_limits:
            locked_cnt = sum(1 for d in range(1, LOCK_UNTIL_DAY + 1) if SPECIFIC_SHIFTS.get((s_idx, d)) == s_type)
            eff_max = max(max_c, locked_cnt)
            if eff_max == 0: model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) == 0)
            else: model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) <= eff_max)

        for s1, s2, _ in ng_pairs:
            for d in range(LOCK_UNTIL_DAY, TARGET_DAYS): model.Add(B[s1, d, S_IRI] + B[s2, d, S_IRI] <= 1)

        for c_idx, (c_name, members, t_shifts, req_count) in enumerate(COMMITTEES):
            Z_days = []
            for d in range(TARGET_DAYS):
                Z_d = model.NewBoolVar(f'comm_{c_idx}_{d}')
                model.Add(sum(B[s, d, t] for s in members for t in t_shifts) >= req_count).OnlyEnforceIf(Z_d)
                Z_days.append(Z_d)
            
            met_any = model.NewBoolVar(f'comm_met_{c_idx}')
            model.AddMaxEquality(met_any, Z_days)
            viol = model.NewBoolVar(f'comm_viol_{c_idx}')
            model.Add(met_any == 1).OnlyEnforceIf(viol.Not())
            model.Add(met_any == 0).OnlyEnforceIf(viol)
            
            obj_terms.append(viol * 50000)

        model.Minimize(sum(obj_terms))
        solver = cp_model.CpSolver()
        solver.parameters.num_search_workers = 8
        solver.parameters.max_time_in_seconds = 15.0
        return solver.Solve(model), solver, B, max_consec, strict_rank

    steps = [
        (4, True, "基本4連勤ルール"), 
        (5, True, "緩和5連勤ルール"), 
        (5, False, "緩和5連勤・順序フリー")
    ]
    used_step = ""
    for c_lim, s_rank, label in steps:
        with st.spinner(f"最小変更で再計算中...（{label}）"):
            status, solver, B, used_limit, used_strict = attempt_reschedule(c_lim, s_rank)
        if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]: 
            used_step = label
            break

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        original_raw_table = [[shift_str[next(t for t in range(10) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)]
        copied_table = [row[:] for row in original_raw_table]
        
        assign_fatal_errors = assign_role_names(copied_table, edited_req_df)
        if assign_fatal_errors:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), assign_fatal_errors, [], pd.DataFrame(), "欠勤フォローモード"
            
        fatal_errs, rule_warnings = validate_assigned_roles(copied_table, original_raw_table)
        if fatal_errs:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), fatal_errs, [], pd.DataFrame(), "欠勤フォローモード"
        if rule_warnings:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), rule_warnings, [], pd.DataFrame(), "欠勤フォローモード"

        if used_limit == 4 and used_strict:
            st.success(f"✅ 最小変更でのシフト再作成が完了しました！")
        else:
            st.warning(f"⚠️ 基本の4連勤ルールでは組めなかったため、一部ルールを緩和（【{used_step}】を適用）して再作成しました。")
            
        result_df = format_result_df(copied_table)

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
        
        with st.expander(f"📝 元のシフト表から変更された箇所の一覧を見る（計 {len(ui_changes)} 箇所）", expanded=True):
            if not ui_changes: st.write("変更された箇所はありません。")
            else:
                for ch in ui_changes: st.markdown(ch)
                
        comm_records = []
        rename_dict = {}
        for c_idx, (c_name, members, t_shifts, req_count) in enumerate(COMMITTEES):
            met_day = None
            attendees = []
            for d in range(TARGET_DAYS):
                cnt = 0
                att_names = []
                for s in members:
                    for t in t_shifts:
                        if solver.Value(B[s, d, t]):
                            cnt += 1
                            actual_shift = copied_table[s][d]
                            att_names.append(f"{staff_names[s]}({actual_shift})")
                if cnt >= req_count:
                    met_day = d + 1
                    attendees = att_names
                    break
            
            if met_day is not None:
                day_str = day_labels[met_day]
                comm_records.append({"委員会名": c_name, "実施日": day_str, "参加対象者": "、".join(attendees)})
                if day_str in rename_dict:
                    rename_dict[day_str] = rename_dict[day_str] + f" ★{c_name}"
                else:
                    rename_dict[day_str] = f"{day_str} ★{c_name}"
            else:
                comm_records.append({"委員会名": c_name, "実施日": "⚠️ 条件を満たす日がありません", "参加対象者": "-"})
        comm_res_df = pd.DataFrame(comm_records)

        result_df = result_df.rename(columns=rename_dict)
        return result_df, changes_df, comm_res_df, [], [], pd.DataFrame(), "欠勤フォローモード"
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
                if req_shift != S_OJT:  model.Add(B[s, d, S_OJT] == 0)
            else:
                if req_shift != S_YUKU and orig_t != S_YUKU: model.Add(B[s, d, S_YUKU] == 0)
                if req_shift != S_KEN and orig_t != S_KEN:  model.Add(B[s, d, S_KEN] == 0)
                if req_shift != S_OJT and orig_t != S_OJT: model.Add(B[s, d, S_OJT] == 0)
                
            if req_shift is not None: model.Add(B[s, d, req_shift] == 1)

    for s in range(staffs): add_soft(model.Add(sum(B[s, d, S_YASU] for d in range(TARGET_DAYS)) == TARGET_YASU), 100000, f"【休数エラー】{staff_names[s]}さんの休が{TARGET_YASU}回ではありません")

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
        C_vars = [model.NewIntVar(0, 100, f'C_{s}_{d}') for d in range(TARGET_DAYS)]
        for d in range(TARGET_DAYS):
            yasu_flag = model.NewBoolVar(f'yasu_{s}_{d}')
            model.Add(B[s, d, S_YASU] + B[s, d, S_YUKU] == 1).OnlyEnforceIf(yasu_flag)
            model.Add(B[s, d, S_YASU] + B[s, d, S_YUKU] == 0).OnlyEnforceIf(yasu_flag.Not())
            
            prev_val = prev_c if d == 0 else C_vars[d-1]
            w_expr = B[s, d, S_HAYA] + B[s, d, S_OSO] + B[s, d, S_Q] + B[s, d, S_KEN] + B[s, d, S_FREE] + B[s, d, S_OJT] + 2 * B[s, d, S_IRI]
            
            model.Add(C_vars[d] == 0).OnlyEnforceIf(yasu_flag)
            model.Add(C_vars[d] == prev_val + w_expr).OnlyEnforceIf(yasu_flag.Not())
            
            viol = model.NewBoolVar(f'viol_{s}_{d}')
            model.Add(C_vars[d] <= 4).OnlyEnforceIf(viol.Not())
            model.Add(C_vars[d] > 4).OnlyEnforceIf(viol)
            
            weight = 15 if (prev_c > 0 and d < 5) else 5
            penalties.append(viol * weight)
            desc_prefix = "月初" if (prev_c > 0 and d < 5) else ""
            violations_info.append((viol, f"【{desc_prefix}連勤超過】{staff_names[s]}さんが{day_labels[d+1]}時点で5連勤以上（基本4連勤ルール違反）"))

    for s_idx, s_type, max_c, desc in personal_limits:
        if max_c == 0: add_soft(model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) == 0), 20, f"【条件違反】{desc}")
        else: add_soft(model.Add(sum(B[s_idx, d, s_type] for d in range(TARGET_DAYS)) <= max_c), 20, f"【条件違反】{desc}")

    for s1, s2, desc in ng_pairs:
        for d in range(TARGET_DAYS): model.Add(B[s1, d, S_IRI] + B[s2, d, S_IRI] <= 1)

    for c_idx, (c_name, members, t_shifts, req_count) in enumerate(COMMITTEES):
        Z_days = []
        for d in range(TARGET_DAYS):
            Z_d = model.NewBoolVar(f'comm_{c_idx}_{d}')
            model.Add(sum(B[s, d, t] for s in members for t in t_shifts) >= req_count).OnlyEnforceIf(Z_d)
            Z_days.append(Z_d)
        
        met_any = model.NewBoolVar(f'comm_met_{c_idx}')
        model.AddMaxEquality(met_any, Z_days)
        viol = model.NewBoolVar(f'comm_viol_{c_idx}')
        model.Add(met_any == 1).OnlyEnforceIf(viol.Not())
        model.Add(met_any == 0).OnlyEnforceIf(viol)
        
        penalties.append(viol * 50000)
        violations_info.append((viol, f"【委員会配置エラー】「{c_name}」の条件を満たす日が1日も作れませんでした"))

    model.Minimize(sum(penalties))
    solver = cp_model.CpSolver()
    solver.parameters.num_search_workers = 8
    solver.parameters.max_time_in_seconds = 15.0
    
    with st.spinner("原因調査モードで計算しています..."): status = solver.Solve(model)
        
    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        broken_rules = [desc for viol_var, desc in violations_info if solver.Value(viol_var)]
            
        original_raw_table = [[shift_str[next(t for t in range(10) if solver.Value(B[s, d, t]))] for d in range(TARGET_DAYS)] for s in range(staffs)]
        copied_table = [row[:] for row in original_raw_table]
        
        assign_fatal_errors = assign_role_names(copied_table, edited_req_df)
        if assign_fatal_errors:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), assign_fatal_errors, [], pd.DataFrame(), "原因調査モード"
            
        fatal_errs, rule_warnings = validate_assigned_roles(copied_table, original_raw_table)
        if fatal_errs:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), fatal_errs, [], pd.DataFrame(), "原因調査モード"
            
        result_df = format_result_df(copied_table)
        
        comm_records = []
        rename_dict = {}
        for c_idx, (c_name, members, t_shifts, req_count) in enumerate(COMMITTEES):
            met_day = None
            attendees = []
            for d in range(TARGET_DAYS):
                cnt = 0
                att_names = []
                for s in members:
                    for t in t_shifts:
                        if solver.Value(B[s, d, t]):
                            cnt += 1
                            actual_shift = copied_table[s][d]
                            att_names.append(f"{staff_names[s]}({actual_shift})")
                if cnt >= req_count:
                    met_day = d + 1
                    attendees = att_names
                    break
            
            if met_day is not None:
                day_str = day_labels[met_day]
                comm_records.append({"委員会名": c_name, "実施日": day_str, "参加対象者": "、".join(attendees)})
                if day_str in rename_dict:
                    rename_dict[day_str] = rename_dict[day_str] + f" ★{c_name}"
                else:
                    rename_dict[day_str] = f"{day_str} ★{c_name}"
            else:
                comm_records.append({"委員会名": c_name, "実施日": "⚠️ 条件を満たす日がありません", "参加対象者": "-"})
        comm_res_df = pd.DataFrame(comm_records)
        result_df = result_df.rename(columns=rename_dict)
        
        # デバッグモード用の診断結果データフレームを作成
        diag_records = []
        for msg in broken_rules + rule_warnings:
            diag_records.append({"対象年月": f"{target_year}年{target_month}月", "実行モード": "原因調査モード", "診断内容": msg})
        if not diag_records:
            diag_records.append({"対象年月": f"{target_year}年{target_month}月", "実行モード": "原因調査モード", "診断内容": "制約違反や未確定の勤務番号は見つかりませんでした。"})
        diag_df = pd.DataFrame(diag_records)
        
        return result_df, pd.DataFrame(), comm_res_df, [], rule_warnings, diag_df, "原因調査モード"
    else:
        st.error("指定自体が矛盾しています。")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), [], [], pd.DataFrame(), "原因調査モード"

# =====================================================================
# 3. 実行ボタンと結果出力
# =====================================================================
st.subheader("3. シフトの作成と出力")
btn_label = "🛠️ 最小変更でシフトを再作成する（欠勤フォロー実行）" if "欠勤フォロー" in mode else "🚀 シフト表を作成する"

current_hash = get_state_hash()
if st.session_state.last_calc_hash and st.session_state.last_calc_hash != current_hash:
    st.session_state.assigned_result = None
    st.session_state.assigned_comm = None
    st.session_state.assigned_changes = None
    st.session_state.assigned_warnings = None
    st.session_state.assigned_diag = None
    st.session_state.last_executed_mode = ""
    st.warning("⚠️ **設定や表の入力内容が変更されました。** 古い計算結果は無効化されました。最新のシフト表を出力するには再計算してください。")

if st.button(btn_label, type="primary", use_container_width=True):
    st.session_state.assigned_result = None
    
    fatal_input_errors = validate_manual_inputs(edited_req_df, staff_names, TARGET_DAYS)
    if fatal_input_errors:
        st.error("🚨 **手動指定に致命的な矛盾が見つかりました。以下のエラーを修正してください。**")
        for e in fatal_input_errors: st.error(f"・ {e}")
    else:
        if check_shortage():
            if "通常作成" in mode: 
                res_tuple = run_normal_mode()
            elif "欠勤フォロー" in mode: 
                res_tuple = run_reschedule_mode()
            else: 
                res_tuple = run_debug_mode()
                
            if len(res_tuple) == 7:
                result_df, changes_df, comm_res_df, assign_fatal_errors, assign_warnings, diag_df, executed_mode = res_tuple
            else:
                result_df, changes_df, comm_res_df, assign_fatal_errors, assign_warnings, diag_df, executed_mode = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), [], [], pd.DataFrame(), ""
                
            if assign_fatal_errors:
                st.error("🚨 **手動指定の矛盾等、表示割り当て中に致命的なエラーが見つかりました。**")
                for e in assign_fatal_errors: st.error(f"・ {e}")
            elif not result_df.empty:
                st.session_state.assigned_result = result_df
                st.session_state.assigned_comm = comm_res_df
                st.session_state.assigned_changes = changes_df
                st.session_state.assigned_warnings = assign_warnings
                st.session_state.assigned_diag = diag_df
                st.session_state.last_executed_mode = executed_mode
                st.session_state.last_calc_hash = get_state_hash()

# キャッシュに計算結果があれば表示する（再描画対応）
if st.session_state.assigned_result is not None:
    result_df = st.session_state.assigned_result
    comm_res_df = st.session_state.assigned_comm
    changes_df = st.session_state.assigned_changes
    assign_warnings = st.session_state.assigned_warnings
    diag_df = st.session_state.assigned_diag
    executed_mode = st.session_state.last_executed_mode

    is_diagnostic = (executed_mode == "原因調査モード")

    if is_diagnostic:
        st.error("⚠️ **【診断用シフト表】** このシフト表は人数不足やルール違反を含んでいます。正式な完成シフトとして使用しないでください。")
        if assign_warnings:
            unique_warnings = list(dict.fromkeys(assign_warnings))
            with st.expander("⚠️ 番号・アルファベットの割り当てに関する警告（シフトの人数不整合等）", expanded=True):
                for w in unique_warnings: st.warning(w)
    
    if not comm_res_df.empty:
        st.markdown("### 📋 AIが決定した委員会・グループ日程")
        st.table(comm_res_df)
        st.divider()

    styled_result = style_shift_df(result_df)
    st.dataframe(styled_result, use_container_width=False, width=2000, height=min(620, 65 + staffs * 35))
    
    result_sheet_name = "診断用シフト表（未確定）" if is_diagnostic else "完成シフト表"
    file_prefix = "diagnostic_shift" if is_diagnostic else "shift_schedule"
    dl_button_label = "📥 【診断用】シフト表をExcelでダウンロード（制約違反あり）" if is_diagnostic else "📥 完成したシフト表をExcelでダウンロード（色分け・回数集計・全設定同梱）"
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        styled_result.to_excel(writer, sheet_name=result_sheet_name)
        if is_diagnostic and not diag_df.empty:
            diag_df.to_excel(writer, sheet_name='診断結果', index=False)
        if not comm_res_df.empty:
            comm_res_df.to_excel(writer, sheet_name='委員会日程表', index=False)
        if not changes_df.empty:
            changes_df.to_excel(writer, sheet_name='変更依頼リスト', index=False)
        settings_df.to_excel(writer, sheet_name='基本設定', index=False)
        edited_staff_df.to_excel(writer, sheet_name='スタッフ設定', index=False)
        edited_ng_pairs_df.to_excel(writer, sheet_name='NGペア設定', index=False)
        edited_committee_df.to_excel(writer, sheet_name='委員会設定', index=False)
        edited_personal_lock_df.to_excel(writer, sheet_name='個別期間ロック設定', index=False)
        save_req_df.to_excel(writer, sheet_name='希望休入力')
        apply_excel_dropdowns(writer, staffs, TARGET_DAYS)
    excel_data = output.getvalue()
    
    st.download_button(
        label=dl_button_label,
        data=excel_data,
        file_name=f"{file_prefix}_{target_year}_{target_month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
