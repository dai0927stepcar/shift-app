"""介護施設向けシフト自動作成。実行: streamlit run shift_app.py

通常勤務と番号付きOJTは異なる状態。1職員1日につき1状態を割り当てる。
UI以外の関数はStreamlitセッションに依存せず、Excelとモデルを個別検証できる。
"""
from __future__ import annotations

import calendar
import datetime as dt
import hashlib
import io
import json
import math
import re
from dataclasses import dataclass, field

import pandas as pd
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from ortools.sat.python import cp_model


PASSWORD = "0927"
NORMAL = "通常作成モード（高速・自動3段階緩和）"
FOLLOW = "🛠️ 欠勤フォロー（最小変更）モード"
DEBUG = "原因調査モード（デバッグ用・強制出力）"
MODES = [NORMAL, FOLLOW, DEBUG]
PREV_COL = "前月連勤"
WEEKDAYS_JP = ["月", "火", "水", "木", "金", "土", "日"]
DEFAULT_Q_DAYS = [1, 3, 6, 7, 8, 13, 14, 15, 17, 20, 22, 24, 27, 28, 29]
S_HAYA, S_Q, S_OSO, S_IRI, S_AKE, S_YASU, S_YUKU, S_KEN, S_FREE, S_OJT, S_BATH = range(11)
SHIFT_CATEGORY = {
    **{r: S_HAYA for r in ["早①", "早②", "早③"]}, "Q": S_Q,
    **{r: S_OSO for r in ["遅①", "遅②", "遅③"]},
    "A夜": S_IRI, "B夜": S_IRI, "A明": S_AKE, "B明": S_AKE,
    "休": S_YASU, "有給": S_YUKU, "研修": S_KEN, "フリー": S_FREE,
    "OJT": S_OJT, "入浴": S_BATH,
}
OJT_TARGETS = ["早①", "早②", "早③", "遅①", "遅②", "遅③", "Q", "A夜", "B夜"]
OJT_RECOVERY = ["A明(OJT)", "B明(OJT)"]
OJT_LABELS = [f"{r}(OJT)" for r in OJT_TARGETS] + OJT_RECOVERY
REGULAR_ROLES = ["早①", "早②", "早③", "Q", "遅①", "遅②", "遅③", "A夜", "B夜", "A明", "B明"]
LABELS = list(SHIFT_CATEGORY) + OJT_LABELS
SHIFT_OPTIONS = [""] + list(SHIFT_CATEGORY) + ["希望休"] + OJT_LABELS
ALLOWED_SHIFTS = set(SHIFT_OPTIONS)
STAFF_COLUMNS = ["スタッフ名", "職員区分", "早の上限", "Qの上限", "遅の上限", "入の上限", "属性1", "属性2", "属性3"]
COMM_SHIFT_COL = "対象シフト（例: 遅①,A明）"
COMM_COLUMNS = ["委員会名", COMM_SHIFT_COL, "コアメンバー"]
OJT_COLUMNS = ["対象職員", "対象勤務区分", "月間OJT実施回数"]
HISTORY_COLUMNS = ["日付", "スタッフ名", "変更前", "変更後", "備考"]
FIXED_OJT_COLUMNS = ["スタッフ名", "日付", "勤務区分"]
COLORS = {"休": "FFD6D6", "希望休": "FFD6D6", "有給": "DCEBFF", "研修": "FFE8CC", "委員会": "F4B183"}


def clean(value):
    if value is None:
        return ""
    if isinstance(value, (list, tuple, set)):
        return value
    return "" if pd.isna(value) else str(value).strip()


def integer(value, name, low=0, high=31):
    try:
        number = float(value)
        if not math.isfinite(number) or not number.is_integer() or not low <= number <= high:
            raise ValueError
        return int(number)
    except (ValueError, TypeError, OverflowError):
        raise ValueError(f"{name}は{low}〜{high}の整数で指定してください（入力: {value}）。") from None


def base_role(label):
    return label.removesuffix("(OJT)")


def category(label):
    return SHIFT_CATEGORY["休" if label == "希望休" else base_role(label)]


def canonical(label):
    return "休" if label == "希望休" else label


def typed_ojt(label):
    return label in OJT_LABELS


def ojt_execution(label):
    return typed_ojt(label) and base_role(label) in OJT_TARGETS


def work_units(label):
    kind = category(label)
    if kind in (S_YASU, S_YUKU, S_AKE):
        return 0
    return 2 if kind == S_IRI else 1


def extract_day_num(column):
    match = re.match(r"^(\d+)(?:\s*\(|\s*$)", str(column).strip())
    return int(match.group(1)) if match else None


def parse_people(value):
    if isinstance(value, (list, tuple, set)):
        return list(dict.fromkeys(str(v).strip() for v in value if str(v).strip()))
    value = clean(value)
    if not value:
        return []
    try:
        parsed = json.loads(value)
        if isinstance(parsed, list) and all(isinstance(v, str) for v in parsed):
            return list(dict.fromkeys(v.strip() for v in parsed if v.strip()))
    except (ValueError, TypeError):
        pass
    return list(dict.fromkeys(v.strip() for v in re.split(r"[,、\n]", value) if v.strip()))


def people_json(value):
    return json.dumps(parse_people(value), ensure_ascii=False)


def empty_df(columns):
    return pd.DataFrame(columns=columns).astype(object)


def get_default_staff_df():
    return pd.DataFrame([
        {"スタッフ名": chr(65 + i), "職員区分": "通常職員", "早の上限": "制限なし",
         "Qの上限": "制限なし", "遅の上限": "制限なし", "入の上限": "制限なし",
         "属性1": "", "属性2": "", "属性3": ""} for i in range(16)
    ])


def normalize_staff(frame):
    frame = frame.copy().fillna("")
    if "スタッフ名" not in frame:
        raise ValueError("スタッフ設定に「スタッフ名」列がありません。")
    for col in STAFF_COLUMNS:
        if col not in frame:
            frame[col] = "通常職員" if col == "職員区分" else ("制限なし" if "上限" in col else "")
    for col in STAFF_COLUMNS:
        frame[col] = frame[col].map(clean)
    names = frame["スタッフ名"].tolist()
    if not names or any(not name for name in names):
        raise ValueError("スタッフ名を空欄にすることはできません。")
    if len(names) != len(set(names)):
        raise ValueError("スタッフ名が重複しています。")
    frame["職員区分"] = frame["職員区分"].replace("", "通常職員")
    if not frame["職員区分"].isin(["通常職員", "例外職員"]).all():
        raise ValueError("職員区分は通常職員・例外職員から選択してください。")
    for col in [c for c in STAFF_COLUMNS if "上限" in c]:
        frame[col] = frame[col].map(lambda v: "制限なし" if v in ("", "制限なし") else str(integer(v, col)))
    return frame[STAFF_COLUMNS].reset_index(drop=True)


@dataclass
class ScheduleData:
    year: int
    month: int
    days: int
    holidays: int
    q_off: list
    staff: pd.DataFrame
    requests: pd.DataFrame
    ng_pairs: pd.DataFrame = field(default_factory=lambda: empty_df(["スタッフ1", "スタッフ2"]))
    personal_locks: pd.DataFrame = field(default_factory=lambda: empty_df(["スタッフ名", "開始日", "終了日"]))
    committees: pd.DataFrame = field(default_factory=lambda: empty_df(COMM_COLUMNS))
    ojt: pd.DataFrame = field(default_factory=lambda: empty_df(OJT_COLUMNS))
    mode: str = NORMAL
    lock_until: int = 0
    base: pd.DataFrame | None = None
    ojt_history: pd.DataFrame = field(default_factory=lambda: empty_df(HISTORY_COLUMNS))
    saved_committees: pd.DataFrame = field(default_factory=pd.DataFrame)
    saved_ojt: pd.DataFrame = field(default_factory=pd.DataFrame)
    fixed_ojt: pd.DataFrame = field(default_factory=lambda: empty_df(FIXED_OJT_COLUMNS))

    @property
    def names(self):
        return self.staff["スタッフ名"].tolist()

    def day_label(self, day):
        return f"{day}({WEEKDAYS_JP[dt.date(self.year, self.month, day).weekday()]})"

    def request(self, name, day):
        return clean(self.requests.at[name, str(day)])

    def original(self, name, day):
        return clean(self.base.at[name, str(day)]) if self.base is not None else ""

    def exceptional(self, name):
        return self.staff.loc[self.staff["スタッフ名"] == name, "職員区分"].iloc[0] == "例外職員"

    def members(self, committee):
        return [row["スタッフ名"] for _, row in self.staff.iterrows()
                if committee in {clean(row[f"属性{i}"]) for i in (1, 2, 3)}]


def default_data():
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=9)))
    days = calendar.monthrange(today.year, today.month)[1]
    staff = get_default_staff_df()
    requests = pd.DataFrame("", index=staff["スタッフ名"], columns=[PREV_COL] + [str(d) for d in range(1, days + 1)])
    requests[PREV_COL] = "0"
    return ScheduleData(today.year, today.month, days, 9, [d for d in DEFAULT_Q_DAYS if d <= days], staff, requests)


def target_categories(value):
    aliases = {"早": S_HAYA, "遅": S_OSO, "入": S_IRI, "明": S_AKE, "A夜勤": S_IRI, "B夜勤": S_IRI}
    result = set()
    for part in re.split(r"[,、・\s]+", clean(value)):
        if not part:
            continue
        if part in aliases:
            result.add(aliases[part])
        elif part in ALLOWED_SHIFTS and part != "入浴":
            result.add(category(part))
        else:
            raise ValueError(f"委員会の対象勤務区分「{part}」は指定できません。")
    if not result:
        raise ValueError("委員会の対象勤務区分が空欄です。")
    return result


def committee_specs(data):
    return [(clean(r["委員会名"]), data.members(clean(r["委員会名"])),
             target_categories(r[COMM_SHIFT_COL]), parse_people(r.get("コアメンバー", "")))
            for _, r in data.committees.iterrows() if clean(r.get("委員会名", ""))]


def ojt_plan(data):
    plan = {}
    for _, row in data.ojt.iterrows():
        if all(not clean(row.get(c, "")) for c in OJT_COLUMNS):
            continue
        key = (clean(row["対象職員"]), clean(row["対象勤務区分"]))
        if key in plan:
            raise ValueError(f"OJT設定が重複しています: {key[0]} / {key[1]}。回数を1行にまとめてください。")
        plan[key] = integer(row["月間OJT実施回数"], "月間OJT実施回数", 0, data.days)
    return plan


def personal_locked(data, name, day):
    for _, row in data.personal_locks.iterrows():
        if clean(row.get("スタッフ名", "")) == name:
            start = integer(row["開始日"], "開始日", 1, 31)
            end = integer(row["終了日"], "終了日", 1, 31)
            if min(start, end) <= day <= max(start, end):
                return True
    return False


def fixed_value(data, name, day):
    """画面・事前検証・モデルで共用する固定セルの判定。"""
    requested = data.request(name, day)
    if data.exceptional(name):
        return requested or "休"
    if data.mode != FOLLOW:
        return requested or None
    for _, row in data.fixed_ojt.iterrows():
        if clean(row["スタッフ名"]) == name and integer(row["日付"], "OJT固定日", 1, data.days) == day and requested == clean(row["勤務区分"]):
            return requested
    original = data.original(name, day)
    if day <= data.lock_until or personal_locked(data, name, day):
        return requested or original or "休"
    if requested and requested != original:
        return requested
    if requested in ["希望休", "休", "有給", "研修", "OJT", "フリー", "入浴"]:
        return requested
    # 番号付きOJTは未変更なら自動調整可。単独OJTの旧入力は従来どおり固定。
    return None


def collect_fixed_ojt(data):
    records = []
    for name in data.names:
        for day in range(1, data.days + 1):
            requested = data.request(name, day)
            if not typed_ojt(requested):
                continue
            old_fixed = any(clean(row["スタッフ名"]) == name and int(row["日付"]) == day and clean(row["勤務区分"]) == requested
                            for _, row in data.fixed_ojt.iterrows())
            if data.mode != FOLLOW or requested != data.original(name, day) or old_fixed:
                records.append({"スタッフ名": name, "日付": day, "勤務区分": requested})
    return pd.DataFrame(records, columns=FIXED_OJT_COLUMNS)


def validate_data(data):
    errors = []
    try:
        normalize_staff(data.staff)
        integer(data.holidays, "休の日数", 1, 15)
        integer(data.lock_until, "過去日ロック", 0, data.days)
        for day in data.q_off:
            integer(day, "Qを配置しない日", 1, data.days)
        if data.mode == FOLLOW and data.base is None:
            errors.append("欠勤フォローには基準となる完成シフト表が必要です。")
        for name in data.names:
            integer(data.requests.at[name, PREV_COL], f"{name}さんの前月連勤", 0, 5)
            for day in range(1, data.days + 1):
                value = data.request(name, day)
                if value not in ALLOWED_SHIFTS:
                    errors.append(f"{name}さんの{day}日の勤務区分「{value}」は不正です。")
                if value == "入浴" and not data.exceptional(name):
                    errors.append(f"{name}さんの{day}日: 入浴は例外職員専用です。")
        for frame, cols in [(data.ng_pairs, ["スタッフ1", "スタッフ2"]),
                            (data.personal_locks, ["スタッフ名"]), (data.ojt, ["対象職員"])]:
            for _, row in frame.iterrows():
                if all(not clean(v) for v in row):
                    continue
                for col in cols:
                    if clean(row.get(col, "")) not in data.names:
                        errors.append(f"{col}「{clean(row.get(col, ''))}」がスタッフ設定に存在しません。")
        for _, row in data.personal_locks.iterrows():
            if any(clean(v) for v in row):
                integer(row["開始日"], "開始日", 1, 31)
                integer(row["終了日"], "終了日", 1, 31)
        seen = set()
        for title, members, targets, cores in committee_specs(data):
            if title in seen:
                errors.append(f"委員会名「{title}」が重複しています。")
            seen.add(title)
            if not members:
                errors.append(f"委員会「{title}」に所属者がいません。属性1〜3を設定してください。")
            invalid = set(cores) - set(members)
            if invalid:
                errors.append(f"委員会「{title}」のコアメンバーが所属者ではありません: {'、'.join(sorted(invalid))}")
        for (name, role), count in ojt_plan(data).items():
            if role not in OJT_TARGETS:
                errors.append(f"{name}さんのOJT対象「{role}」は指定できません。夜勤明け単独も不可です。")
        for _, row in data.fixed_ojt.iterrows():
            if clean(row["スタッフ名"]) not in data.names or clean(row["勤務区分"]) not in OJT_LABELS:
                errors.append("OJT固定指定のスタッフ名・勤務区分が不正です。")
            integer(row["日付"], "OJT固定日", 1, data.days)
    except (ValueError, KeyError, TypeError) as exc:
        errors.append(str(exc))
    return list(dict.fromkeys(errors))


def validate_manual_inputs(data):
    """通常担当者の重複のみ禁止。通常枠とOJT枠の同番号は許可する。"""
    errors = []
    for day in range(1, data.days + 1):
        fixed = [canonical(v) for name in data.names if (v := fixed_value(data, name, day))]
        for role in REGULAR_ROLES:
            if fixed.count(role) > 1:
                errors.append(f"{day}日: 通常勤務「{role}」の固定指定が重複しています。")
    for name in data.names:
        for day in range(1, data.days):
            current = fixed_value(data, name, day)
            following = fixed_value(data, name, day + 1)
            if current and category(current) == S_IRI and following:
                expected = current.replace("夜", "明")
                if canonical(following) != expected:
                    errors.append(f"{name}さんの{day}日「{current}」と翌日「{following}」が矛盾しています。必要: {expected}")
    return errors


@dataclass
class SolveResult:
    schedule: pd.DataFrame = field(default_factory=pd.DataFrame)
    committees: pd.DataFrame = field(default_factory=pd.DataFrame)
    changes: pd.DataFrame = field(default_factory=lambda: empty_df(HISTORY_COLUMNS))
    ojt_summary: pd.DataFrame = field(default_factory=pd.DataFrame)
    ojt_history: pd.DataFrame = field(default_factory=lambda: empty_df(HISTORY_COLUMNS))
    diagnostics: list = field(default_factory=list)
    diagnostic: bool = False
    step: str = ""
    status: str = ""
    attempts: list = field(default_factory=list)


def build_model(data, max_consec=4, strict_rank=True, diagnostic=False):
    model = cp_model.CpModel()
    names, days = data.names, range(data.days)
    x = {(name, day, label): model.NewBoolVar(f"x_{s}_{day}_{i}")
         for s, name in enumerate(names) for day in days for i, label in enumerate(LABELS)}
    plan = ojt_plan(data)
    violations, diagnostic_terms = [], []
    rest_terms, normal_terms, ojt_terms, five_terms = [], [], [], []

    def total(name, day, kind):
        return sum(x[name, day, label] for label in LABELS if category(label) == kind)

    def soft(constraint, weight, description):
        violation = model.NewBoolVar(f"violation_{len(violations)}")
        constraint.OnlyEnforceIf(violation.Not())
        diagnostic_terms.append(violation * weight)
        violations.append((violation, description))

    for name in names:
        for day in days:
            model.AddExactlyOne(x[name, day, label] for label in LABELS)
            requested, original = data.request(name, day + 1), data.original(name, day + 1)
            fixed = fixed_value(data, name, day + 1)
            if fixed:
                model.Add(x[name, day, canonical(fixed)] == 1)
            for label in ["有給", "研修", "OJT"]:
                if canonical(requested) != label and not (data.mode == FOLLOW and original == label):
                    model.Add(x[name, day, label] == 0)
            if not (data.exceptional(name) and requested == "入浴"):
                model.Add(x[name, day, "入浴"] == 0)
            for label in OJT_LABELS:
                role = base_role(label)
                allowed = (name, role) in plan and plan[name, role] > 0
                if label in OJT_RECOVERY:
                    allowed = (name, role.replace("明", "夜")) in plan
                    if day > 0:
                        entering = label.replace("明", "夜")
                        allowed |= data.request(name, day) == entering or (data.mode == FOLLOW and data.original(name, day) == entering)
                    # 月初のOJT明けは前月からの手入力・基準表のみ。
                    if day == 0:
                        allowed = False
                allowed |= requested == label or (data.mode == FOLLOW and original == label)
                if not allowed:
                    model.Add(x[name, day, label] == 0)

        if not data.exceptional(name):
            rest_count = sum(x[name, day, "休"] for day in days)
            if diagnostic:
                soft(model.Add(rest_count == data.holidays), 100000, f"【休数エラー】{name}さんの休が{data.holidays}回ではありません。")
            elif data.mode == FOLLOW:
                shortage = model.NewIntVar(0, data.days, f"holiday_shortage_{name}")
                model.Add(rest_count + shortage == data.holidays)
                rest_terms.append(shortage)
            else:
                model.Add(rest_count == data.holidays)

    for day in days:
        if data.mode == FOLLOW and day < data.lock_until and not diagnostic:
            continue
        for role in REGULAR_ROLES:
            expected = 0 if role == "Q" and day + 1 in data.q_off else 1
            constraint = model.Add(sum(x[name, day, role] for name in names) == expected)
            if diagnostic:
                soft(constraint, 20, f"【人数エラー】{data.day_label(day + 1)}の通常勤務「{role}」が{expected}人ではありません。OJTは通常人数に含みません。")
        if day + 1 in data.q_off:
            constraint = model.Add(sum(x[name, day, "Q(OJT)"] for name in names) == 0)
            if diagnostic:
                soft(constraint, 20, f"【Q配置エラー】{data.day_label(day + 1)}はQの配置禁止日です（OJTを含む）。")

    start = max(0, data.lock_until - 1) if data.mode == FOLLOW and not diagnostic else 0
    for name in names:
        for day in range(start, data.days - 1):
            for entering, recovery in [("A夜", "A明"), ("B夜", "B明"), ("A夜(OJT)", "A明(OJT)"), ("B夜(OJT)", "B明(OJT)")]:
                constraint = model.Add(x[name, day, entering] == x[name, day + 1, recovery])
                if diagnostic:
                    soft(constraint, 50, f"【リレー違反】{name}さんの{day + 1}日{entering}と翌日の{recovery}が連動していません。")
            constraint = model.Add(x[name, day + 1, "休"] + x[name, day + 1, "有給"] >= total(name, day, S_AKE))
            if diagnostic:
                soft(constraint, 50, f"【リレー違反】{name}さんの{day + 1}日「明」の次が休・有給ではありません。")
            if strict_rank or diagnostic:
                for before, after in [(S_Q, S_HAYA), (S_OSO, S_HAYA), (S_OSO, S_Q)]:
                    constraint = model.Add(total(name, day, before) + total(name, day + 1, after) <= 1)
                    if diagnostic:
                        labels = {S_Q: "Q", S_OSO: "遅", S_HAYA: "早"}
                        soft(constraint, 10, f"【順位違反】{name}さんの{day + 1}日から翌日の勤務順序違反（{labels[before]}→{labels[after]}）。")

        previous = integer(data.requests.at[name, PREV_COL], "前月連勤", 0, 5)
        for day in days:
            consecutive = model.NewIntVar(0, 100, f"consecutive_{name}_{day}")
            rest = model.NewBoolVar(f"rest_{name}_{day}")
            model.Add(rest == x[name, day, "休"] + x[name, day, "有給"])
            model.Add(consecutive == 0).OnlyEnforceIf(rest)
            units = sum(work_units(label) * x[name, day, label] for label in LABELS)
            model.Add(consecutive == previous + units).OnlyEnforceIf(rest.Not())
            if diagnostic:
                soft(model.Add(consecutive <= 4), 15 if day < 5 and int(data.requests.at[name, PREV_COL]) else 5,
                     f"【連勤超過】{name}さんが{day + 1}日時点で基本4連勤を超えています（OJT・入浴も勤務）。")
            elif data.mode != FOLLOW or day >= data.lock_until:
                model.Add(consecutive <= max_consec)
            if max_consec == 5 or data.mode == FOLLOW:
                excess = model.NewIntVar(0, 100, f"five_{name}_{day}")
                model.AddMaxEquality(excess, [0, consecutive - 4])
                five_terms.append(excess)
            previous = consecutive

    for _, row in data.staff.iterrows():
        name = row["スタッフ名"]
        for col, kind in [("早の上限", S_HAYA), ("Qの上限", S_Q), ("遅の上限", S_OSO), ("入の上限", S_IRI)]:
            if row[col] == "制限なし":
                continue
            maximum = integer(row[col], col)
            if data.mode == FOLLOW and not diagnostic:
                locked_count = sum(1 for d in range(1, data.lock_until + 1)
                                   if (v := fixed_value(data, name, d)) and category(v) == kind)
                maximum = max(maximum, locked_count)
            constraint = model.Add(sum(total(name, d, kind) for d in days) <= maximum)
            if diagnostic:
                soft(constraint, 20, f"【条件違反】{name}さんの{col}（OJTを含む）が{maximum}回を超えています。")

    for _, row in data.ng_pairs.iterrows():
        first, second = clean(row["スタッフ1"]), clean(row["スタッフ2"])
        if not first and not second:
            continue
        if first == second:
            continue
        for day in days:
            if data.mode == FOLLOW and day < data.lock_until and not diagnostic:
                continue
            model.Add(total(first, day, S_IRI) + total(second, day, S_IRI) <= 1)

    committee_days = {}
    for index, (title, members, targets, cores) in enumerate(committee_specs(data)):
        selected = [model.NewBoolVar(f"committee_{index}_{day}") for day in days]
        model.AddExactlyOne(selected)
        committee_days[title] = selected
        eligible = {label for label in LABELS if category(label) in targets and label != "入浴"}
        majority = len(members) // 2 + 1
        if diagnostic:
            core_bad = model.NewBoolVar(f"core_bad_{index}")
            majority_bad = model.NewBoolVar(f"majority_bad_{index}")
            diagnostic_terms += [core_bad * 50000, majority_bad * 50000]
            violations += [(core_bad, f"【委員会コア出席エラー】「{title}」のコア全員（{'、'.join(cores)}）の対象勤務を、手動指定・夜勤連動等と両立できません。"),
                           (majority_bad, f"【委員会過半数エラー】「{title}」の所属{len(members)}人中{majority}人以上が対象勤務になる開催日を確保できません。")]
        for day in days:
            attendance = {name: sum(x[name, day, label] for label in eligible) for name in members}
            condition = [selected[day], majority_bad.Not()] if diagnostic else selected[day]
            model.Add(sum(attendance.values()) >= majority).OnlyEnforceIf(condition)
            for name in cores:
                condition = [selected[day], core_bad.Not()] if diagnostic else selected[day]
                model.Add(attendance[name] == 1).OnlyEnforceIf(condition)

    for (name, role), target in plan.items():
        count = sum(x[name, day, f"{role}(OJT)"] for day in days)
        if diagnostic:
            soft(model.Add(count == target), 2000, f"【OJT回数エラー】{name}さんの{role}(OJT)を予定{target}回配置できません。手動勤務・例外職員の勤務可能日・上限・連勤・委員会条件を確認してください。")
        elif data.mode == FOLLOW:
            deficit = model.NewIntVar(0, data.days, f"ojt_deficit_{name}_{role}")
            model.AddMaxEquality(deficit, [0, target - count])
            fixed_count = sum(fixed_value(data, name, d) == f"{role}(OJT)" for d in range(1, data.days + 1))
            model.Add(count <= max(target, fixed_count))
            ojt_terms.append(deficit * 10)
        else:
            model.Add(count == target)

    if diagnostic:
        model.Minimize(sum(diagnostic_terms))
    elif data.mode == FOLLOW:
        for name in names:
            for day in days:
                original = data.original(name, day + 1)
                if not original or day < data.lock_until:
                    continue
                old = canonical(original)
                if typed_ojt(old):
                    ojt_terms.append((1 - x[name, day, old]) * 5)
                else:
                    kind = category(old)
                    cost = 5 if kind in (S_OJT, S_FREE) else 20 if kind in (S_IRI, S_AKE) else 30 if kind == S_YASU else 10
                    same_category = sum(x[name, day, label] for label in LABELS
                                        if not typed_ojt(label) and category(label) == kind)
                    normal_terms += [(1 - same_category) * cost, 1 - x[name, day, old]]
        normal_terms.extend(term * 50 for term in five_terms)
        # 下位目的の最大値を上回る係数で、通常勤務の維持をOJTより厳密に優先。
        ojt_bound = 15 * len(names) * data.days + 10 * sum(plan.values()) + 1
        normal_bound = 5050 * len(names) * data.days + 1
        model.Minimize(sum(rest_terms) * normal_bound * ojt_bound + sum(normal_terms) * ojt_bound + sum(ojt_terms))
    else:
        iri_counts = []
        for name in names:
            count = model.NewIntVar(0, data.days, f"iri_count_{name}")
            model.Add(count == sum(x[name, day, role] for day in days for role in ["A夜", "B夜"]))
            iri_counts.append(count)
        maximum = model.NewIntVar(0, data.days, "max_iri")
        model.AddMaxEquality(maximum, iri_counts)
        model.Minimize(maximum * 100 + sum(five_terms))
    return model, x, violations, committee_days


def format_result_df(data, table):
    frame = pd.DataFrame(table, index=data.names, columns=[data.day_label(d) for d in range(1, data.days + 1)])
    for name in data.names:
        for day in range(1, data.days + 1):
            if data.request(name, day) == "希望休" and frame.at[name, data.day_label(day)] == "休":
                frame.at[name, data.day_label(day)] = "希望休"
    day_cols = list(frame.columns)
    for title, values in [("【早】", ["早①", "早②", "早③"]), ("【Q】", ["Q"]),
                          ("【遅】", ["遅①", "遅②", "遅③"]), ("【入】", ["A夜", "B夜"]),
                          ("【明】", ["A明", "B明"]), ("【休計】", ["休", "希望休"]),
                          ("【有給】", ["有給"]), ("【研修】", ["研修"]),
                          ("【OJT】", [f"{r}(OJT)" for r in OJT_TARGETS] + ["OJT"]),
                          ("【OJT明】", OJT_RECOVERY), ("【フリー】", ["フリー"]), ("【入浴】", ["入浴"])]:
        frame[title] = frame[day_cols].isin(values).sum(axis=1)
    return frame


def inspect_schedule(data, frame, limit=4, strict=True, include_past=False, check_plans=True):
    """求解モデルとは別に完成表の人数・固定・リレー・OJT等を検証する。"""
    frame = frame.rename(columns={column: data.day_label(day) for column in frame.columns
                                  if (day := extract_day_num(column)) is not None})
    errors = []
    start = 1 if include_past or data.mode != FOLLOW else data.lock_until + 1
    for day in range(start, data.days + 1):
        values = [canonical(frame.at[name, data.day_label(day)]) for name in data.names]
        for role in REGULAR_ROLES:
            expected = 0 if role == "Q" and day in data.q_off else 1
            if values.count(role) != expected:
                errors.append(f"{day}日の通常勤務{role}: {values.count(role)}名（必要{expected}名）。")
    for name in data.names:
        consecutive = integer(data.requests.at[name, PREV_COL], "前月連勤", 0, 5)
        rests = 0
        for day in range(1, data.days + 1):
            value = canonical(frame.at[name, data.day_label(day)])
            fixed = fixed_value(data, name, day)
            if fixed and value != canonical(fixed):
                errors.append(f"{name}さん{day}日の固定勤務{fixed}が守られていません。")
            if value == "入浴" and not (data.exceptional(name) and data.request(name, day) == "入浴"):
                errors.append(f"{name}さん{day}日の入浴が手動指定と一致しません。")
            rests += value == "休"
            consecutive = 0 if category(value) in (S_YASU, S_YUKU) else consecutive + work_units(value)
            if day >= start and consecutive > limit:
                errors.append(f"{name}さん{day}日の連勤数{consecutive}が{limit}を超えています。")
            if day < data.days and day >= max(1, start - 1):
                following = canonical(frame.at[name, data.day_label(day + 1)])
                if category(value) == S_IRI and following != value.replace("夜", "明"):
                    errors.append(f"{name}さん{day}日からの夜勤連動が不正です。")
                if category(following) == S_AKE and value != following.replace("明", "夜"):
                    errors.append(f"{name}さん{day + 1}日の明けの前日が不正です。")
                if category(value) == S_AKE and category(following) not in (S_YASU, S_YUKU):
                    errors.append(f"{name}さん{day}日明けの翌日が休・有給ではありません。")
                if strict and (category(value), category(following)) in [(S_Q, S_HAYA), (S_OSO, S_HAYA), (S_OSO, S_Q)]:
                    errors.append(f"{name}さん{day}日から翌日の勤務順序が不正です。")
        if not data.exceptional(name) and ((data.mode != FOLLOW and rests != data.holidays) or (data.mode == FOLLOW and rests > data.holidays)):
            errors.append(f"{name}さんの休数{rests}が設定と一致しません。")
        row = data.staff.loc[data.staff["スタッフ名"] == name].iloc[0]
        for col, kind in [("早の上限", S_HAYA), ("Qの上限", S_Q), ("遅の上限", S_OSO), ("入の上限", S_IRI)]:
            if row[col] == "制限なし":
                continue
            count = sum(category(frame.at[name, data.day_label(d)]) == kind for d in range(1, data.days + 1))
            maximum = integer(row[col], col)
            if data.mode == FOLLOW:
                maximum = max(maximum, sum(category(frame.at[name, data.day_label(d)]) == kind for d in range(1, data.lock_until + 1)))
            if count > maximum:
                errors.append(f"{name}さんの{col}を超えています。")
    for _, row in data.ng_pairs.iterrows():
        first, second = clean(row["スタッフ1"]), clean(row["スタッフ2"])
        if first and second and first != second:
            for day in range(start, data.days + 1):
                if category(frame.at[first, data.day_label(day)]) == category(frame.at[second, data.day_label(day)]) == S_IRI:
                    errors.append(f"{day}日の{first}さんと{second}さんの夜勤NGペア違反です。")
    for title, members, targets, cores in committee_specs(data):
        if not any(all(category(frame.at[n, data.day_label(d)]) in targets for n in cores)
                   and sum(category(frame.at[n, data.day_label(d)]) in targets for n in members) >= len(members) // 2 + 1
                   for d in range(1, data.days + 1)):
            errors.append(f"委員会「{title}」のコア全員出席と過半数出席を満たす日がありません。")
    if check_plans:
        for (name, role), target in ojt_plan(data).items():
            count = sum(frame.at[name, data.day_label(d)] == f"{role}(OJT)" for d in range(1, data.days + 1))
            if data.mode != FOLLOW and count != target:
                errors.append(f"{name}さんの{role}(OJT)の回数{count}が予定{target}と一致しません。")
    return list(dict.fromkeys(errors))


def make_ojt_reports(data, result):
    keys = set(ojt_plan(data))
    for name in data.names:
        for day in range(1, data.days + 1):
            for value in [data.original(name, day), result.schedule.at[name, data.day_label(day)]]:
                if ojt_execution(value):
                    keys.add((name, base_role(value)))
    records, history = [], []
    for name, role in sorted(keys):
        target = ojt_plan(data).get((name, role), 0)
        original_count = sum(data.original(name, d) == f"{role}(OJT)" for d in range(1, data.days + 1))
        manual_count = sum(data.request(name, d) == f"{role}(OJT)" for d in range(1, data.days + 1))
        planned = max(target, original_count, manual_count)
        if not data.saved_ojt.empty and {"スタッフ名", "対象勤務区分", "予定回数"}.issubset(data.saved_ojt.columns):
            saved = data.saved_ojt[(data.saved_ojt["スタッフ名"] == name) & (data.saved_ojt["対象勤務区分"] == role)]
            if not saved.empty and int(saved.iloc[0].get("設定回数", target)) == target:
                planned = max(planned, integer(saved.iloc[0]["予定回数"], "保存済みOJT予定回数", 0, data.days))
        actual = sum(result.schedule.at[name, data.day_label(d)] == f"{role}(OJT)" for d in range(1, data.days + 1))
        records.append({"スタッフ名": name, "対象勤務区分": role, "設定回数": target, "予定回数": planned,
                        "変更後回数": actual, "中止回数": max(0, planned - actual)})
    if data.base is not None:
        for name in data.names:
            for day in range(1, data.days + 1):
                original = data.original(name, day)
                actual = result.schedule.at[name, data.day_label(day)]
                if ojt_execution(original) and original != actual:
                    history.append({"日付": data.day_label(day), "スタッフ名": name, "変更前": original, "変更後": actual,
                                    "備考": "手動変更" if data.request(name, day) != original else "自動調整（OJT変更・中止）"})
    result.ojt_summary = pd.DataFrame(records, columns=["スタッフ名", "対象勤務区分", "設定回数", "予定回数", "変更後回数", "中止回数"])
    result.ojt_history = pd.concat([data.ojt_history, pd.DataFrame(history, columns=HISTORY_COLUMNS)], ignore_index=True).drop_duplicates().reset_index(drop=True)


def solve_once(data, max_consec=4, strict_rank=True, diagnostic=False, seconds=15.0, workers=8):
    model, variables, violations, committee_days = build_model(data, max_consec, strict_rank, diagnostic)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = seconds
    solver.parameters.num_search_workers = workers
    solver.parameters.random_seed = 927
    status = solver.Solve(model)
    result = SolveResult(diagnostic=diagnostic, status=solver.StatusName(status))
    if status not in (cp_model.FEASIBLE, cp_model.OPTIMAL):
        result.diagnostics = ["制限時間内に解を確認できませんでした。" if status == cp_model.UNKNOWN else
                              "固定勤務・例外職員の未指定日・夜勤連動・NGペア等の必須条件が両立しません。"]
        return result
    table = [[next(label for label in LABELS if solver.Value(variables[name, day, label])) for day in range(data.days)] for name in data.names]
    result.schedule = format_result_df(data, table)
    result.diagnostics = [description for variable, description in violations if solver.Value(variable)]
    committee_records, rename = [], {}
    for title, members, targets, cores in committee_specs(data):
        day = next(d + 1 for d, selected in enumerate(committee_days[title]) if solver.Value(selected))
        attending = [n for n in members if category(result.schedule.at[n, data.day_label(day)]) in targets]
        valid = set(cores).issubset(attending) and len(attending) >= len(members) // 2 + 1
        committee_records.append({"委員会名": title, "実施日": data.day_label(day) if valid else "⚠️ 条件を満たす日がありません",
                                  "参加対象者": "、".join(f"{n}({result.schedule.at[n, data.day_label(day)]})" for n in attending) if valid else "-",
                                  "参加職員JSON": people_json(attending if valid else [])})
        if valid:
            key = data.day_label(day)
            rename[key] = rename.get(key, key) + f" ★{title}"
    result.committees = pd.DataFrame(committee_records, columns=["委員会名", "実施日", "参加対象者", "参加職員JSON"])
    independent = inspect_schedule(data, result.schedule, max_consec, strict_rank, check_plans=not diagnostic)
    if independent and not diagnostic:
        result.schedule = pd.DataFrame()
        result.diagnostics = ["完成表の検証エラー: " + message for message in independent]
        return result
    if diagnostic:
        result.diagnostics = list(dict.fromkeys(result.diagnostics + independent))
    changes = []
    if data.base is not None:
        for name in data.names:
            for day in range(1, data.days + 1):
                old, new = data.original(name, day), result.schedule.at[name, data.day_label(day)]
                if old and old != new:
                    changes.append({"日付": data.day_label(day), "スタッフ名": name, "変更前": old, "変更後": new,
                                    "備考": "手動変更（欠勤・指定）" if data.request(name, day) != old else "自動調整（フォロー）"})
    result.changes = pd.DataFrame(changes, columns=HISTORY_COLUMNS)
    make_ojt_reports(data, result)
    result.schedule = result.schedule.rename(columns=rename)
    return result


def run_schedule(data, seconds=15.0, workers=8, progress=None):
    errors = validate_data(data)
    if errors:
        return SolveResult(diagnostics=errors, status="INPUT_ERROR")
    if data.mode == DEBUG:
        result = solve_once(data, diagnostic=True, seconds=seconds, workers=workers)
        result.step = "原因調査モード"
        return result
    errors = validate_manual_inputs(data)
    if errors:
        return SolveResult(diagnostics=errors, status="INPUT_ERROR")
    attempts = []
    for limit, strict, title in [(4, True, "基本4連勤ルール"), (5, True, "緩和5連勤ルール"), (5, False, "緩和5連勤・順序フリー")]:
        if progress:
            progress(title)
        result = solve_once(data, limit, strict, seconds=seconds, workers=workers)
        attempts.append(f"{title}: {result.status}")
        if not result.schedule.empty:
            result.step, result.attempts = title, attempts
            return result
    if progress:
        progress("原因調査モード")
    diagnostic = solve_once(data, diagnostic=True, seconds=seconds, workers=workers)
    diagnostic.step, diagnostic.attempts = "原因調査モード", attempts
    if any("UNKNOWN" in attempt for attempt in attempts):
        diagnostic.diagnostics.insert(0, "通常計算に時間切れを含みます。条件が不可能と証明されたわけではありません。")
    return diagnostic


def normalize_grid(frame, names, days, title, require_prev=False, complete=False):
    frame = frame.copy().fillna("")
    frame.index = frame.index.map(lambda value: str(value).strip())
    if frame.index.has_duplicates or set(frame.index) != set(names):
        raise ValueError(f"「{title}」のスタッフ名をスタッフ設定と一致させてください。重複も不可です。")
    result = pd.DataFrame("", index=names, columns=[PREV_COL] + [str(d) for d in range(1, days + 1)])
    result[PREV_COL] = "0"
    if require_prev and PREV_COL not in frame:
        raise ValueError(f"「{title}」に前月連勤列がありません。")
    if PREV_COL in frame:
        for name in names:
            result.at[name, PREV_COL] = str(integer(frame.at[name, PREV_COL], f"{name}さんの前月連勤", 0, 5))
    day_columns = {}
    for column in frame.columns:
        day = extract_day_num(column)
        if day and 1 <= day <= days:
            if day in day_columns:
                raise ValueError(f"「{title}」の{day}日列が重複しています。")
            day_columns[day] = column
    for day in range(1, days + 1):
        if day not in day_columns:
            raise ValueError(f"「{title}」に{day}日列がありません。")
        for name in names:
            value = clean(frame.at[name, day_columns[day]])
            if value not in ALLOWED_SHIFTS or (complete and not value):
                raise ValueError(f"「{title}」の{name}さん{day}日: 勤務区分「{value}」が不正、または完成表が空欄です。")
            result.at[name, str(day)] = value
    return result


def read_excel(file, is_base_schedule=False):
    """検証がすべて完了するまで、呼び出し側の状態を更新しない。旧形式も補完する。"""
    if hasattr(file, "seek"):
        file.seek(0)
    with pd.ExcelFile(file, engine="openpyxl") as book:
        required = {"基本設定", "スタッフ設定", "希望休入力"}
        if is_base_schedule:
            required.add("完成シフト表")
        missing = required - set(book.sheet_names)
        if missing:
            raise ValueError("必要なシートがありません: " + "、".join(sorted(missing)))
        settings_frame = pd.read_excel(book, "基本設定").fillna("")
        if not {"設定項目", "設定値"}.issubset(settings_frame.columns):
            raise ValueError("基本設定の必須列（設定項目、設定値）がありません。")
        settings = dict(zip(settings_frame["設定項目"], settings_frame["設定値"]))
        year = integer(settings.get("対象年"), "対象年", 2024, 2035)
        month = integer(settings.get("対象月"), "対象月", 1, 12)
        days = calendar.monthrange(year, month)[1]
        if integer(settings.get("月の日数"), "月の日数", 28, 31) != days:
            raise ValueError("対象年月と月の日数が一致しません。")
        holidays = integer(settings.get("休の日数"), "休の日数", 1, 15)
        raw_q = clean(settings.get("Qを配置しない日", ""))
        q_off = [] if raw_q in ("", "なし") else [integer(v, "Qを配置しない日", 1, days) for v in raw_q.split(",")]
        staff = normalize_staff(pd.read_excel(book, "スタッフ設定"))
        names = staff["スタッフ名"].tolist()
        requests = normalize_grid(pd.read_excel(book, "希望休入力", index_col=0), names, days, "希望休入力", require_prev=True)

        def optional(sheet, columns):
            if sheet not in book.sheet_names:
                return empty_df(columns)
            frame = pd.read_excel(book, sheet).fillna("")
            if not set(columns).issubset(frame.columns):
                # コア指定のない既存ファイルは空のコア指定で補完。
                if sheet == "委員会設定" and set(columns) - set(frame.columns) == {"コアメンバー"}:
                    frame["コアメンバー"] = ""
                else:
                    raise ValueError(f"「{sheet}」に必須列がありません: {'、'.join(sorted(set(columns) - set(frame.columns)))}")
            return frame[columns].astype(object)

        ng = optional("NGペア設定", ["スタッフ1", "スタッフ2"])
        locks = optional("個別期間ロック設定", ["スタッフ名", "開始日", "終了日"])
        for column in ["開始日", "終了日"]:
            locks[column] = locks[column].map(lambda v: "" if not clean(v) else str(integer(v, column, 1, 31)))
        committees = optional("委員会設定", COMM_COLUMNS)
        committees["コアメンバー"] = committees["コアメンバー"].map(people_json)
        ojt = optional("OJT設定", OJT_COLUMNS)
        if not ojt.empty:
            ojt["月間OJT実施回数"] = ojt["月間OJT実施回数"].map(lambda v: str(integer(v, "月間OJT実施回数", 0, days)))
        mode, lock_until, base = NORMAL, 0, None
        if "実行状態" in book.sheet_names:
            state_frame = pd.read_excel(book, "実行状態").fillna("")
            if not {"設定項目", "設定値"}.issubset(state_frame.columns):
                raise ValueError("実行状態シートの必須列がありません。")
            state = dict(zip(state_frame["設定項目"], state_frame["設定値"]))
            mode = clean(state.get("実行モード", NORMAL))
            if mode not in MODES:
                raise ValueError("保存された実行モードが不正です。")
            lock_until = integer(state.get("過去日ロック", 0), "過去日ロック", 0, days)
            if "基準シフト表" in book.sheet_names:
                base = normalize_grid(pd.read_excel(book, "基準シフト表", index_col=0), names, days, "基準シフト表", complete=True)
        fixed_ojt = optional("OJT固定指定", FIXED_OJT_COLUMNS)
        if "OJT固定指定" not in book.sheet_names and mode != FOLLOW:
            fixed_ojt = pd.DataFrame([
                {"スタッフ名": name, "日付": day, "勤務区分": requests.at[name, str(day)]}
                for name in names for day in range(1, days + 1) if typed_ojt(requests.at[name, str(day)])
            ], columns=FIXED_OJT_COLUMNS)
        if is_base_schedule:
            base = normalize_grid(pd.read_excel(book, "完成シフト表", index_col=0), names, days, "完成シフト表", complete=True)
            base[PREV_COL] = requests[PREV_COL]
            requests = base.copy()
            mode, lock_until = FOLLOW, 0
        if base is not None:
            base[PREV_COL] = requests[PREV_COL]
        data = ScheduleData(year, month, days, holidays, sorted(set(q_off)), staff, requests, ng, locks, committees, ojt, mode, lock_until, base)
        data.ojt_history = optional("OJT変更・中止履歴", HISTORY_COLUMNS)
        data.saved_committees = pd.read_excel(book, "委員会日程表").fillna("") if "委員会日程表" in book.sheet_names else pd.DataFrame()
        data.saved_ojt = pd.read_excel(book, "OJT実施結果").fillna("") if "OJT実施結果" in book.sheet_names else pd.DataFrame()
        data.fixed_ojt = fixed_ojt
        for column in ["設定回数", "予定回数", "変更後回数", "中止回数"]:
            if column in data.saved_ojt:
                data.saved_ojt[column] = data.saved_ojt[column].map(lambda v: integer(v, f"OJT実施結果の{column}", 0, days))
        errors = validate_data(data)
        if errors:
            raise ValueError("\n".join(errors))
        return data


def verify_and_load_excel(file, is_base_schedule=False):
    try:
        return [], read_excel(file, is_base_schedule)
    except (ValueError, KeyError, TypeError, OSError, io.UnsupportedOperation) as exc:
        return [str(exc)], None
    except Exception as exc:
        return [f"Excelを読み込めませんでした。ファイル形式を確認してください: {type(exc).__name__}"], None


def add_a3_print_sheet(writer, result_df, data, comm_res_df, changes_df):
    wb = writer.book
    ws = wb.create_sheet("掲示用A3", 0)
    wb.active = ws
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.3
    ws.print_options.horizontalCentered = True
    ws.freeze_panes = "B6"
    ws.print_title_rows = "1:5"
    thin = Side(style="thin", color="000000")
    normal_border = Border(top=thin, bottom=thin, left=thin, right=thin)
    red = Side(style="medium", color="FF0000")
    changed_border = Border(top=red, bottom=red, left=red, right=red)
    changed = {(clean(row["スタッフ名"]), extract_day_num(row["日付"])) for _, row in changes_df.iterrows()}
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=data.days + 1)
    ws.cell(1, 1, f"{data.year}年{data.month}月 勤務表").font = Font(name="Meiryo", size=16, bold=True)
    ws.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
    valid_committees = []
    for _, row in comm_res_df.iterrows():
        if extract_day_num(row["実施日"]):
            valid_committees.append(f"{row['委員会名']}: {row['実施日']}")
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=data.days + 1)
    ws.cell(3, 1, "★ 委員会実施日： " + " ／ ".join(valid_committees) if valid_committees else "")
    ws.cell(3, 1).font = Font(name="Meiryo", size=10, bold=True, color="0000FF")
    ws.cell(3, 1).alignment = Alignment(horizontal="right", shrink_to_fit=True)
    ws.merge_cells("A4:A5")
    ws.cell(4, 1, "職員名").border = normal_border
    ws.cell(4, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 18
    day_cols = [c for c in result_df if extract_day_num(c)]
    for index, column in enumerate(day_cols, 2):
        day = extract_day_num(column)
        weekday = dt.date(data.year, data.month, day).weekday()
        for row, value in [(4, f"★\n{day}" if "★" in str(column) else day), (5, WEEKDAYS_JP[weekday])]:
            cell = ws.cell(row, index, value)
            cell.border = normal_border
            cell.font = Font(name="Meiryo", size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if weekday in (5, 6):
                cell.fill = PatternFill("solid", fgColor="DCE6FA" if weekday == 5 else "F7DADA")
        ws.column_dimensions[get_column_letter(index)].width = 6
    ws.row_dimensions[4].height = 28
    for index, name in enumerate(data.names):
        row = index + 6
        ws.row_dimensions[row].height = 26
        ws.cell(row, 1, name).font = Font(name="Meiryo", size=10)
        ws.cell(row, 1).border = normal_border
        for column_index, column in enumerate(day_cols, 2):
            source = f"'完成シフト表'!{get_column_letter(column_index)}{index + 2}"
            cell = ws.cell(row, column_index, f'=IF({source}="","",{source})')
            cell.font = Font(name="Meiryo", size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, shrink_to_fit=True)
            cell.border = changed_border if (name, extract_day_num(column)) in changed else normal_border
    area = f"B6:{get_column_letter(data.days + 1)}{len(data.names) + 5}"
    for value in ["休", "希望休", "有給", "研修"]:
        ws.conditional_formatting.add(area, CellIsRule(operator="equal", formula=[f'"{value}"'],
                                                      fill=PatternFill("solid", fgColor=COLORS[value]), stopIfTrue=True))
    specs = {title: (members, targets) for title, members, targets, cores in committee_specs(data)}
    # 所属者の開催日セルにだけ、現在の勤務が参加条件を満たす場合の条件付き書式を付与。
    # 休・希望休・有給・研修は委員会色の対象から除外し、通常色を優先する。
    for _, record in comm_res_df.iterrows():
        day = extract_day_num(record["実施日"])
        if not day or record["委員会名"] not in specs:
            continue
        members, targets = specs[record["委員会名"]]
        eligible = [label for label in LABELS if category(label) in targets and label not in ["休", "希望休", "有給", "研修", "入浴"]]
        for name in members:
            address = f"{get_column_letter(day + 1)}{data.names.index(name) + 6}"
            if eligible:
                formula = "OR(" + ",".join(f'{address}="{label}"' for label in eligible) + ")"
                ws.conditional_formatting.add(address, FormulaRule(formula=[formula], fill=PatternFill("solid", fgColor=COLORS["委員会"]), stopIfTrue=True))
    legend_row = len(data.names) + 7
    ws.cell(legend_row, 1, "【凡例】").font = Font(name="Meiryo", size=9, bold=True)
    for index, value in enumerate(["休・希望休", "有給", "研修", "委員会"]):
        column = 2 + index * 3
        ws.merge_cells(start_row=legend_row, start_column=column, end_row=legend_row, end_column=column + 2)
        cell = ws.cell(legend_row, column, value)
        cell.fill = PatternFill("solid", fgColor=COLORS["休" if index == 0 else value])
        cell.font = Font(name="Meiryo", size=9)
        cell.alignment = Alignment(horizontal="center", shrink_to_fit=True)
    ws.print_area = f"A1:{get_column_letter(data.days + 1)}{legend_row}"


def apply_excel_dropdowns(writer, data):
    # Excelの直接列挙リストは255文字まで。OJT拡張後は非表示シート＋名前付き範囲を使う。
    from openpyxl.workbook.defined_name import DefinedName
    wb = writer.book
    options = wb.create_sheet("入力候補")
    normal_options = [v for v in SHIFT_OPTIONS if v and v != "入浴"]
    all_options = [v for v in SHIFT_OPTIONS if v]
    for column, values in [(1, normal_options), (2, all_options)]:
        for row, value in enumerate(values, 1):
            options.cell(row, column, value)
    for title, column, length in [("NormalShiftOptions", "A", len(normal_options)), ("AllShiftOptions", "B", len(all_options))]:
        definition = DefinedName(title, attr_text=f"'入力候補'!${column}$1:${column}${length}")
        if hasattr(wb.defined_names, "add"):
            wb.defined_names.add(definition)
        else:
            wb.defined_names.append(definition)
    options.sheet_state = "hidden"
    for title, offset in [("希望休入力", 3), ("基準シフト表", 3), ("完成シフト表", 2), ("診断用シフト表（未確定）", 2)]:
        if title not in writer.sheets:
            continue
        ws = writer.sheets[title]
        for row, name in enumerate(data.names, 2):
            dv = DataValidation(type="list", formula1="=AllShiftOptions" if data.exceptional(name) else "=NormalShiftOptions",
                                allow_blank=True, showErrorMessage=True, errorStyle="stop", errorTitle="入力エラー",
                                error="リストから勤務を選択してください。入浴は例外職員専用です。")
            ws.add_data_validation(dv)
            dv.add(f"{get_column_letter(offset)}{row}:{get_column_letter(offset + data.days - 1)}{row}")
        if offset == 3:
            dv = DataValidation(type="list", formula1='"0,1,2,3,4,5"', allow_blank=False, showErrorMessage=True, errorStyle="stop")
            ws.add_data_validation(dv)
            dv.add(f"B2:B{len(data.names) + 1}")


def style_shift_df(frame):
    def style(value):
        if value in COLORS and value != "委員会":
            return f"background-color: #{COLORS[value]};"
        return ""
    return frame.style.map(style) if hasattr(frame.style, "map") else frame.style.applymap(style)


def export_excel(data, result=None):
    output = io.BytesIO()
    settings = pd.DataFrame([
        {"設定項目": "対象年", "設定値": data.year}, {"設定項目": "対象月", "設定値": data.month},
        {"設定項目": "月の日数", "設定値": data.days}, {"設定項目": "休の日数", "設定値": data.holidays},
        {"設定項目": "Qを配置しない日", "設定値": ",".join(map(str, sorted(data.q_off))) or "なし"},
    ])
    requests = data.requests.copy()
    requests.columns = [PREV_COL] + [data.day_label(d) for d in range(1, data.days + 1)]
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if result is not None and not result.schedule.empty:
            sheet = "診断用シフト表（未確定）" if result.diagnostic else "完成シフト表"
            style_shift_df(result.schedule).to_excel(writer, sheet_name=sheet)
            if not result.diagnostic:
                add_a3_print_sheet(writer, result.schedule, data, result.committees, result.changes)
            if result.diagnostic:
                pd.DataFrame({"対象年月": [f"{data.year}年{data.month}月"] * len(result.diagnostics),
                              "実行モード": ["原因調査モード"] * len(result.diagnostics), "診断内容": result.diagnostics}).to_excel(writer, sheet_name="診断結果", index=False)
            result.committees.to_excel(writer, sheet_name="委員会日程表", index=False)
            result.ojt_summary.to_excel(writer, sheet_name="OJT実施結果", index=False)
            result.ojt_history.to_excel(writer, sheet_name="OJT変更・中止履歴", index=False)
            if not result.changes.empty:
                result.changes.to_excel(writer, sheet_name="変更依頼リスト", index=False)
        else:
            if not data.saved_committees.empty:
                data.saved_committees.to_excel(writer, sheet_name="委員会日程表", index=False)
            if not data.saved_ojt.empty:
                data.saved_ojt.to_excel(writer, sheet_name="OJT実施結果", index=False)
            data.ojt_history.to_excel(writer, sheet_name="OJT変更・中止履歴", index=False)
        settings.to_excel(writer, sheet_name="基本設定", index=False)
        data.staff.to_excel(writer, sheet_name="スタッフ設定", index=False)
        data.ng_pairs.to_excel(writer, sheet_name="NGペア設定", index=False)
        data.personal_locks.to_excel(writer, sheet_name="個別期間ロック設定", index=False)
        data.committees.to_excel(writer, sheet_name="委員会設定", index=False)
        data.ojt.to_excel(writer, sheet_name="OJT設定", index=False)
        collect_fixed_ojt(data).to_excel(writer, sheet_name="OJT固定指定", index=False)
        requests.to_excel(writer, sheet_name="希望休入力")
        pd.DataFrame([{"設定項目": "実行モード", "設定値": data.mode},
                      {"設定項目": "過去日ロック", "設定値": data.lock_until}]).to_excel(writer, sheet_name="実行状態", index=False)
        if data.base is not None:
            data.base.to_excel(writer, sheet_name="基準シフト表")
        apply_excel_dropdowns(writer, data)
    return output.getvalue()


def get_state_hash(data):
    def frame(value):
        return None if value is None else value.fillna("").astype(str).to_json(force_ascii=False, orient="split")
    state = {"year": data.year, "month": data.month, "days": data.days, "holidays": data.holidays,
             "q_off": sorted(data.q_off), "mode": data.mode, "lock": data.lock_until,
             **{key: frame(getattr(data, key)) for key in ["staff", "requests", "ng_pairs", "personal_locks", "committees", "ojt", "base", "ojt_history", "fixed_ojt"]}}
    return hashlib.sha256(json.dumps(state, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def sync_staff(data, edited):
    raw = edited.copy()
    edited = normalize_staff(edited)
    old = data.names
    new = edited["スタッフ名"].tolist()
    # data_editorは削除行のindexを保持する。行の対応から改名を関連設定にも反映する。
    mapping = {}
    for index, row in raw.iterrows():
        if index in data.staff.index:
            mapping[data.staff.at[index, "スタッフ名"]] = clean(row["スタッフ名"])
    requests = data.requests.loc[list(mapping)].rename(index=mapping).reindex(new, fill_value="").copy()
    requests[PREV_COL] = requests[PREV_COL].replace("", "0")
    data.requests = requests
    if data.base is not None:
        data.base = data.base.rename(index=mapping).reindex(new) if set(mapping) == set(old) and len(old) == len(new) else None
    for frame, columns in [(data.ng_pairs, ["スタッフ1", "スタッフ2"]),
                           (data.personal_locks, ["スタッフ名"]), (data.ojt, ["対象職員"]), (data.fixed_ojt, ["スタッフ名"])]:
        for column in columns:
            frame[column] = frame[column].replace(mapping)
    if not data.committees.empty:
        data.committees["コアメンバー"] = data.committees["コアメンバー"].map(lambda v: people_json([mapping.get(n, n) for n in parse_people(v)]))
    data.staff = edited


def resize_month(data, year, month):
    days = calendar.monthrange(year, month)[1]
    data.requests = data.requests.reindex(columns=[PREV_COL] + [str(d) for d in range(1, days + 1)], fill_value="")
    data.year, data.month, data.days = year, month, days
    data.q_off = [d for d in data.q_off if d <= days]
    data.base, data.lock_until = None, 0
    data.saved_committees, data.saved_ojt = pd.DataFrame(), pd.DataFrame()
    data.ojt_history = empty_df(HISTORY_COLUMNS)
    data.fixed_ojt = empty_df(FIXED_OJT_COLUMNS)


def main():
    import streamlit as st

    st.set_page_config(page_title="シフト自動作成システム", layout="wide")
    st.title("📅 シフト自動作成システム")
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        st.subheader("ログイン")
        with st.form("login"):
            password = st.text_input("🔒 パスワードを入力してください", type="password")
            submitted = st.form_submit_button("ログイン", type="primary")
        if submitted:
            if password == PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("パスワードが違います")
        st.stop()
    if "data" not in st.session_state:
        st.session_state.data = default_data()
    if "editor_key" not in st.session_state:
        st.session_state.editor_key = 0
    if "result" not in st.session_state:
        st.session_state.result = None
    if "last_calc_hash" not in st.session_state:
        st.session_state.last_calc_hash = ""
    data = st.session_state.data
    ek = st.session_state.editor_key
    if st.session_state.pop("input_applied", False):
        st.success("入力内容を反映しました。")

    def reset_editors():
        st.session_state.editor_key += 1
        st.session_state.result = None
        st.rerun()

    st.sidebar.header("⚙️ 毎月の基本設定")
    col_year, col_month = st.sidebar.columns(2)
    year = int(col_year.number_input("対象年", 2024, 2035, data.year, key=f"year_{ek}"))
    month = int(col_month.number_input("対象月", 1, 12, data.month, key=f"month_{ek}"))
    if (year, month) != (data.year, data.month):
        resize_month(data, year, month)
        reset_editors()
    st.sidebar.info(f"月の日数: {data.days}日（カレンダー自動計算）")
    data.holidays = int(st.sidebar.number_input("通常職員の「休＋希望休」の日数", 1, 15, data.holidays, key=f"holidays_{ek}"))
    data.q_off = st.sidebar.multiselect("「Q」シフトを【配置しない】日", list(range(1, data.days + 1)),
                                       default=data.q_off, format_func=data.day_label, key=f"q_{ek}")
    data.mode = st.sidebar.radio("実行モード", MODES, index=MODES.index(data.mode), key=f"mode_{ek}")
    if data.mode == FOLLOW:
        st.sidebar.info("完成シフト表を読み込み、欠勤箇所を編集して再作成します。OJTは通常勤務の補充を優先して調整します。")
        data.lock_until = int(st.sidebar.number_input("🔒 何日目までを過去日として固定するか", 0, data.days, data.lock_until, key=f"lock_{ek}"))
    else:
        data.lock_until = 0
    st.sidebar.divider()
    st.sidebar.subheader("📂 保存データの読み込み")
    uploaded = st.sidebar.file_uploader("Excelファイルを選択", type=["xlsx"])
    if uploaded is not None:
        restore = st.sidebar.button("📥 入力途中のデータを復元", use_container_width=True)
        follow = st.sidebar.button("🛠️ 完成シフト表を読み込む（欠勤フォロー用）", type="primary", use_container_width=True)
        if restore or follow:
            errors, restored = verify_and_load_excel(uploaded, is_base_schedule=follow)
            if errors:
                st.sidebar.error("\n".join(errors))
            else:
                st.session_state.data = restored
                st.session_state.last_calc_hash = ""
                reset_editors()

    st.subheader("1. スタッフ名・個人条件の設定")
    with st.expander("👥 スタッフ名・職員区分・勤務上限・属性1〜3", expanded=False):
        st.caption("例外職員は手入力した勤務だけを配置し、未指定日は休にします。共通の休日日数からは除外します。")
        with st.popover("📝 スタッフ名をまとめて一括変更"):
            with st.form(f"bulk_{ek}"):
                names_text = st.text_area("スタッフ名リスト（改行・カンマ区切り）")
                if st.form_submit_button("名前を一括反映する"):
                    proposed = [v.strip() for v in re.split(r"[,、\n]+", names_text) if v.strip()]
                    if not proposed:
                        st.error("スタッフ名を入力してください。")
                    else:
                        staff = data.staff.copy()
                        for index, name in enumerate(proposed):
                            if index < len(staff):
                                staff.at[index, "スタッフ名"] = name
                            else:
                                row = get_default_staff_df().iloc[0].copy()
                                row["スタッフ名"] = name
                                staff = pd.concat([staff, pd.DataFrame([row])], ignore_index=True)
                        try:
                            sync_staff(data, staff)
                            reset_editors()
                        except ValueError as exc:
                            st.error(str(exc))
        with st.form(f"staff_form_{ek}"):
            staff_config = {
                "スタッフ名": st.column_config.TextColumn("スタッフ名", required=True),
                "職員区分": st.column_config.SelectboxColumn("職員区分", options=["通常職員", "例外職員"], default="通常職員", required=True),
                **{col: st.column_config.SelectboxColumn(col + "（0＝禁止）", options=["制限なし"] + [str(i) for i in range(32)], default="制限なし", required=True)
                   for col in ["早の上限", "Qの上限", "遅の上限", "入の上限"]},
                **{f"属性{i}": st.column_config.TextColumn(f"属性{i}（委員会等）", default="") for i in (1, 2, 3)},
            }
            edited_staff = st.data_editor(data.staff, column_config=staff_config, num_rows="dynamic", use_container_width=True, key=f"staff_{ek}")
            if st.form_submit_button("スタッフ設定を反映"):
                try:
                    sync_staff(data, edited_staff)
                    reset_editors()
                except ValueError as exc:
                    st.error(str(exc))
    with st.expander("🚫 夜勤（OJTを含む）の同日NGペア設定"):
        with st.form(f"ng_form_{ek}"):
            edited_ng = st.data_editor(data.ng_pairs, num_rows="dynamic", use_container_width=True, key=f"ng_{ek}", column_config={
                col: st.column_config.SelectboxColumn(col, options=data.names, required=True) for col in ["スタッフ1", "スタッフ2"]})
            if st.form_submit_button("NGペア設定を反映"):
                data.ng_pairs = edited_ng.fillna("")
                reset_editors()
    with st.expander("📋 委員会・グループ設定とコアメンバー"):
        st.caption("対象勤務はカンマ区切りです。同じ勤務区分の番号はまとめて数えます。開催日はコア全員と所属者の過半数の出席を必須とします。")
        with st.form(f"committee_form_{ek}"):
            definitions = st.data_editor(data.committees[["委員会名", COMM_SHIFT_COL]], num_rows="dynamic", use_container_width=True, key=f"comm_{ek}", column_config={
                "委員会名": st.column_config.TextColumn("委員会名", required=True),
                COMM_SHIFT_COL: st.column_config.TextColumn("対象勤務区分（例: 遅①,A明）", required=True)})
            if st.form_submit_button("委員会の名称・対象勤務を反映"):
                old_cores = dict(zip(data.committees["委員会名"], data.committees["コアメンバー"]))
                definitions = definitions.fillna("")
                definitions["コアメンバー"] = definitions["委員会名"].map(lambda n: old_cores.get(n, "[]"))
                data.committees = definitions
                reset_editors()
        if not data.committees.empty:
            with st.form(f"core_form_{ek}"):
                cores = {}
                invalid = {}
                for index, row in data.committees.iterrows():
                    title = clean(row["委員会名"])
                    if not title:
                        continue
                    members = data.members(title)
                    saved = parse_people(row["コアメンバー"])
                    invalid[index] = [n for n in saved if n not in members]
                    if invalid[index]:
                        st.error(f"{title}: 所属から外れたコア指定があります: {'、'.join(invalid[index])}。所属を戻すか、下の解除を選択してください。")
                    cores[index] = st.multiselect(f"{title}のコアメンバー", options=members,
                                                 default=[n for n in saved if n in members], key=f"cores_{ek}_{index}")
                clear_invalid = st.checkbox("所属から外れたコア指定を解除する", key=f"clear_core_{ek}")
                if st.form_submit_button("コアメンバーを反映"):
                    for index, values in cores.items():
                        data.committees.at[index, "コアメンバー"] = people_json(values + ([] if clear_invalid else invalid[index]))
                    reset_editors()
    with st.expander("🎓 職員別・勤務区分別OJT設定"):
        st.caption("OJTは通常配置に追加する人員です。夜勤OJTの翌日は同じA/Bの明け(OJT)になります。例外職員のOJTは、該当日を番号付きOJTで手入力してください。")
        with st.form(f"ojt_form_{ek}"):
            edited_ojt = st.data_editor(data.ojt, num_rows="dynamic", use_container_width=True, key=f"ojt_{ek}", column_config={
                "対象職員": st.column_config.SelectboxColumn("OJT対象職員", options=data.names, required=True),
                "対象勤務区分": st.column_config.SelectboxColumn("対象勤務区分", options=OJT_TARGETS, required=True),
                "月間OJT実施回数": st.column_config.SelectboxColumn("月間OJT実施回数", options=[str(i) for i in range(data.days + 1)], required=True)})
            if st.form_submit_button("OJT設定を反映"):
                data.ojt = edited_ojt.fillna("")
                reset_editors()
    with st.expander("🔒 欠勤フォロー用：個別期間ロック設定"):
        st.caption("対象期間は、画面で指定した勤務を固定します。番号付きOJTも固定対象です。")
        with st.form(f"lock_form_{ek}"):
            locks = st.data_editor(data.personal_locks, num_rows="dynamic", use_container_width=True, key=f"plock_{ek}", column_config={
                "スタッフ名": st.column_config.SelectboxColumn("スタッフ名", options=data.names, required=True),
                **{c: st.column_config.SelectboxColumn(c, options=[str(d) for d in range(1, 32)], required=True) for c in ["開始日", "終了日"]}})
            if st.form_submit_button("個別期間ロックを反映"):
                data.personal_locks = locks.fillna("")
                reset_editors()

    follow_mode = data.mode == FOLLOW
    st.subheader(f"2. {'欠勤箇所の編集' if follow_mode else '希望休・指定シフトの入力'}（{data.year}年{data.month}月）")
    if follow_mode and data.base is None:
        st.error("基準となる完成シフト表を左のメニューから読み込んでください。")
    st.caption("表の編集後は、この下の「入力を反映」または「シフト表を作成」を押してください。入浴は例外職員の表で選択できます。")
    with st.form(f"request_form_{ek}"):
        updated_requests = data.requests.copy()
        for is_exception in [False, True]:
            subset = [n for n in data.names if data.exceptional(n) == is_exception]
            if not subset:
                continue
            if is_exception:
                st.markdown("**例外職員（空欄の日は休・入浴を選択可）**")
            options = SHIFT_OPTIONS if is_exception else [v for v in SHIFT_OPTIONS if v != "入浴"]
            configs = {PREV_COL: st.column_config.SelectboxColumn(PREV_COL, options=[str(i) for i in range(6)], required=True, width="small")}
            configs.update({str(d): st.column_config.SelectboxColumn(("🔒" if follow_mode and d <= data.lock_until else "") + data.day_label(d),
                                                                    options=options, width="small") for d in range(1, data.days + 1)})
            edited = st.data_editor(data.requests.loc[subset], column_config=configs, use_container_width=True,
                                    height=min(620, 65 + len(subset) * 35), key=f"req_{ek}_{is_exception}")
            updated_requests.loc[subset] = edited.fillna("")
        st.subheader("3. シフトの作成と出力")
        buttons = st.columns(2)
        apply_requests = buttons[0].form_submit_button("入力を反映", use_container_width=True)
        execute = buttons[1].form_submit_button("🛠️ 最小変更でシフトを再作成" if follow_mode else "🚀 シフト表を作成", type="primary", use_container_width=True)
    if apply_requests or execute:
        data.requests = updated_requests
        st.session_state.result = None
    current_hash = get_state_hash(data)
    if st.session_state.result is not None and st.session_state.last_calc_hash != current_hash:
        st.session_state.result = None
        st.warning("設定・入力が変更されたため、古い計算結果を無効にしました。再作成してください。")
    if execute:
        placeholder = st.empty()
        with st.spinner("シフトを計算しています…"):
            result = run_schedule(data, progress=lambda title: placeholder.info(f"計算中: {title}"))
        placeholder.empty()
        st.session_state.result = result
        st.session_state.last_calc_hash = get_state_hash(data)
        st.session_state.editor_key += 1
        st.rerun()
    if apply_requests:
        st.session_state.input_applied = True
        st.session_state.editor_key += 1
        st.rerun()
    errors = validate_data(data)
    if errors:
        for error in errors:
            st.error(error)
    else:
        progress_bytes = export_excel(data)
        st.download_button("💾 現在の入力進捗をExcelに一時保存", progress_bytes,
                           file_name=f"shift_progress_{data.year}_{data.month:02d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.sidebar.download_button("💾 現在の入力進捗を保存", progress_bytes,
                                   file_name=f"shift_progress_{data.year}_{data.month:02d}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if follow_mode and st.button("🧹 表を空欄に戻して通常作成モードへ"):
        data.requests.loc[:, [str(d) for d in range(1, data.days + 1)]] = ""
        data.base, data.mode, data.lock_until = None, NORMAL, 0
        data.fixed_ojt = empty_df(FIXED_OJT_COLUMNS)
        reset_editors()
    result = st.session_state.result
    if result is not None:
        if result.diagnostic:
            st.error("⚠️ 診断用シフト表です。制約違反を含む可能性があり、正式な完成シフトとして使用できません。")
        elif not result.schedule.empty:
            if result.step == "基本4連勤ルール":
                st.success("✅ シフト作成完了（基本4連勤ルール）")
            else:
                st.warning(f"一部ルールを緩和して作成しました: {result.step}")
        if result.diagnostics:
            with st.expander("🔍 診断結果・エラー", expanded=True):
                for message in result.diagnostics:
                    st.warning(message)
        if result.attempts:
            with st.expander("計算状況"):
                for attempt in result.attempts:
                    st.text(attempt)
        if not result.schedule.empty:
            if not result.committees.empty:
                st.markdown("**📋 委員会・グループ日程**")
                st.dataframe(result.committees[["委員会名", "実施日", "参加対象者"]], hide_index=True, use_container_width=True)
            if not result.ojt_summary.empty:
                st.markdown("**🎓 OJTの実施回数・中止回数**")
                st.dataframe(result.ojt_summary, hide_index=True, use_container_width=True)
            if not result.ojt_history.empty:
                with st.expander("OJT変更・中止履歴", expanded=True):
                    st.dataframe(result.ojt_history, hide_index=True, use_container_width=True)
            if not result.changes.empty:
                with st.expander(f"📝 元のシフトからの変更一覧（{len(result.changes)}箇所）", expanded=True):
                    st.dataframe(result.changes, hide_index=True, use_container_width=True)
            st.dataframe(style_shift_df(result.schedule), use_container_width=True, height=min(620, 65 + len(data.names) * 35))
            prefix = "diagnostic_shift" if result.diagnostic else "shift_schedule"
            st.download_button("📥 【診断用】シフト表をExcelでダウンロード" if result.diagnostic else "📥 完成シフト表をExcelでダウンロード（A3掲示用・全設定同梱）",
                               export_excel(data, result), file_name=f"{prefix}_{data.year}_{data.month:02d}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.sidebar.divider()
    if st.sidebar.button("🔓 ログアウト"):
        st.session_state.authenticated = False
        st.rerun()


if __name__ == "__main__":
    main()
