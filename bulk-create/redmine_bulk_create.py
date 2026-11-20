#!/usr/bin/env python3
"""
Redmine チケット一括作成ツール

CSV または Excel(.xlsx)に書いたチケット(親子孫の階層つき)を、
Redmine の REST API で作成する。
標準ライブラリだけで動くので、pip でのインストールは不要。

使い方:
  # 試し実行(作成はしない。内容のチェックと、作成予定の一覧を表示)
  python3 redmine_bulk_create.py tickets.csv --project my-project
  python3 redmine_bulk_create.py tickets.xlsx --project my-project

  # 本番実行(実際に作成する)
  python3 redmine_bulk_create.py tickets.xlsx --project my-project --apply

APIキー:
  環境変数 REDMINE_API_KEY、または ~/.redmine_api_key ファイルに書いておく。
  (Redmine の「個人設定」→「APIアクセスキー」で確認できる)
"""

import argparse
import csv
import datetime
import json
import os
import re
import sys
import urllib.error
import urllib.request

# ============================================================
# 設定
# ============================================================

DEFAULT_URL = os.environ.get("REDMINE_URL", "http://localhost:3000")

# 題名の区切り(親の題名 + 区切り + 子の題名)
SUBJECT_SEPARATOR = ":"

# 説明が空欄のときに入れる雛形(トラッカー名ごと)。不要なら {} にする
DESCRIPTION_TEMPLATES = {
    "タスク": "# このタスクのゴール\n\n\n\n# 対応内容まとめ\n\n\n\n",
}

# 担当者欄でこれらを書くと(または空欄なら)、自分が担当者になる
SELF_WORDS = {"", "自分", "me", "<<自分>>"}

# CSV の見出し(日本語・英語のどちらでも書ける)
COLUMN_ALIASES = {
    "key": ["行ID", "ID", "key"],
    "parent": ["親", "親ID", "parent"],
    "tracker": ["トラッカー", "tracker"],
    "subject": ["題名", "タイトル", "subject"],
    "assignee": ["担当者", "assignee"],
    "start_date": ["開始日", "start_date"],
    "due_date": ["期日", "due_date"],
    "priority": ["優先度", "priority"],
    "status": ["ステータス", "status"],
    "version": ["対象バージョン", "バージョン", "version"],
    "category": ["カテゴリ", "category"],
    "estimated_hours": ["予定工数", "estimated_hours"],
    "description": ["説明", "description"],
}
REQUIRED_COLUMNS = ["key", "tracker", "subject"]


# ============================================================
# Redmine API
# ============================================================

class RedmineError(Exception):
    pass


class Redmine:
    def __init__(self, url, api_key):
        self.url = url.rstrip("/")
        self.api_key = api_key

    def _request(self, method, path, body=None):
        data = json.dumps(body).encode("utf-8") if body is not None else None
        req = urllib.request.Request(self.url + path, data=data, method=method)
        req.add_header("X-Redmine-API-Key", self.api_key)
        req.add_header("Content-Type", "application/json")
        try:
            with urllib.request.urlopen(req, timeout=30) as res:
                text = res.read().decode("utf-8")
                return json.loads(text) if text.strip() else {}
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "replace")
            try:
                errors = json.loads(detail).get("errors")
                if errors:
                    detail = " / ".join(errors)
            except ValueError:
                detail = detail[:300]
            raise RedmineError(f"{method} {path} → HTTP {e.code}: {detail}") from None
        except urllib.error.URLError as e:
            raise RedmineError(f"{self.url} に接続できません: {e.reason}") from None

    def get(self, path):
        return self._request("GET", path)

    def post(self, path, body):
        return self._request("POST", path, body)


# ============================================================
# ファイルの読み込み(CSV / Excel)
# ============================================================

# Excel の既定のシート名。このシートがあれば優先して読む(なければ先頭のシート)
DEFAULT_SHEET = "チケット"


def read_csv_table(path):
    """CSV を読み、表(行のリスト)を返す。UTF-8(BOM付き含む)と Shift_JIS に対応"""
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp932"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise SystemExit("CSV の文字コードを判別できません(UTF-8 か Shift_JIS で保存してください)")
    return list(csv.reader(text.splitlines()))


# --- Excel(.xlsx)----------------------------------------------
# 追加のインストールなしで動かすため、標準ライブラリだけで .xlsx を読む。
# (.xlsx の中身は、XML ファイルを zip でまとめたもの)

_NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# 日付として扱う Excel 標準の表示形式の番号(日本語版の「yyyy年m月d日」などを含む)
_DATE_FORMAT_IDS = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))


def _is_date_format(code):
    """独自の表示形式が日付かどうか(引用符や [] の中を除いて y/m/d を含むか)"""
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code or "")
    return bool(re.search(r"[ymd]", code, re.IGNORECASE)) and "General" not in code


def _col_index(ref):
    """セル番地(例: "C12")から、0始まりの列番号を返す"""
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_xlsx_table(path, sheet_name=None):
    """Excel を読み、表(行のリスト)を返す。日付のセルは YYYY-MM-DD の文字列にする"""
    import zipfile
    import xml.etree.ElementTree as ET

    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise SystemExit("Excel ファイルを開けません(.xlsx 形式で保存してください。古い .xls は非対応です)")

    def xml(name):
        return ET.fromstring(zf.read(name))

    # シートの一覧と、ファイル内の場所
    wb = xml("xl/workbook.xml")
    rels = {r.get("Id"): r.get("Target") for r in xml("xl/_rels/workbook.xml.rels").findall("rel:Relationship", _NS)}
    sheets = [(s.get("name"), rels[s.get("{%s}id" % _NS["r"])]) for s in wb.find("m:sheets", _NS)]
    names = [n for n, _ in sheets]
    if sheet_name:
        if sheet_name not in names:
            raise SystemExit(f"シート「{sheet_name}」がありません。あるシート: {'、'.join(names)}")
        target = dict(sheets)[sheet_name]
    else:
        target = dict(sheets).get(DEFAULT_SHEET, sheets[0][1])
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target

    pr = wb.find("m:workbookPr", _NS)
    epoch = datetime.date(1904, 1, 1) if pr is not None and pr.get("date1904") in ("1", "true") else datetime.date(1899, 12, 30)

    # 共有文字列(セルの文字は、ここにまとめて保存されている)
    shared = []
    if "xl/sharedStrings.xml" in zf.namelist():
        for si in xml("xl/sharedStrings.xml").findall("m:si", _NS):
            shared.append("".join(t.text or "" for t in si.iter("{%s}t" % _NS["m"])))

    # 表示形式(日付のセルを見分けるため)
    date_styles = set()
    if "xl/styles.xml" in zf.namelist():
        st = xml("xl/styles.xml")
        custom = {}
        fmts = st.find("m:numFmts", _NS)
        if fmts is not None:
            custom = {int(f.get("numFmtId")): f.get("formatCode") for f in fmts}
        xfs = st.find("m:cellXfs", _NS)
        if xfs is not None:
            for i, xf in enumerate(xfs):
                fid = int(xf.get("numFmtId", 0))
                if fid in _DATE_FORMAT_IDS or (fid in custom and _is_date_format(custom[fid])):
                    date_styles.add(i)

    table = {}
    for row in xml(target).iter("{%s}row" % _NS["m"]):
        r = int(row.get("r")) - 1
        for c in row.findall("m:c", _NS):
            typ, style = c.get("t"), int(c.get("s", 0))
            v = c.find("m:v", _NS)
            raw = v.text if v is not None else None
            if typ == "s":
                val = shared[int(raw)] if raw is not None else ""
            elif typ == "inlineStr":
                val = "".join(t.text or "" for t in c.iter("{%s}t" % _NS["m"]))
            elif typ in ("str", "e"):
                val = raw or ""
            elif typ == "b":
                val = "TRUE" if raw == "1" else "FALSE"
            elif typ == "d":
                val = (raw or "")[:10]
            elif raw is None:
                val = ""
            else:
                num = float(raw)
                if style in date_styles:
                    val = (epoch + datetime.timedelta(days=int(num))).isoformat()
                elif num.is_integer():
                    val = str(int(num))
                else:
                    val = str(num)
            table.setdefault(r, {})[_col_index(c.get("r"))] = val

    if not table:
        return []
    width = max(max(cols) for cols in table.values()) + 1
    return [[table.get(r, {}).get(c, "") for c in range(width)] for r in range(max(table) + 1)]


def read_rows(path, sheet_name=None):
    """CSV または Excel を読み、チケットの行(列名 → 値)のリストを返す"""
    if path.lower().endswith((".xlsx", ".xlsm")):
        table, kind = read_xlsx_table(path, sheet_name), "Excel"
    else:
        table, kind = read_csv_table(path), "CSV"
    if not table:
        return []

    header_map = {}
    for idx, name in enumerate(table[0]):
        clean = (name or "").strip()
        for key, aliases in COLUMN_ALIASES.items():
            if clean in aliases:
                header_map[idx] = key
    missing = [k for k in REQUIRED_COLUMNS if k not in header_map.values()]
    if missing:
        names = ", ".join(COLUMN_ALIASES[k][0] for k in missing)
        raise SystemExit(f"{kind} の1行目(見出し)に、必須の列がありません: {names}")

    rows = []
    for line_no, rec in enumerate(table[1:], start=2):
        row = {key: (rec[idx] if idx < len(rec) else "").strip() for idx, key in header_map.items()}
        if not any(row.values()):
            continue  # 空行は無視
        if row.get("key", "").startswith("#") and not row.get("tracker"):
            continue  # 「#」で始まるメモ行は無視
        row["_line"] = line_no
        rows.append(row)
    return rows


# ============================================================
# 名前 → ID の解決
# ============================================================

def normalize_date(value):
    if not value:
        return None
    m = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", value)
    if not m:
        raise ValueError(f"日付の形式が不正です: {value}(例: 2027-01-31)")
    y, mo, d = map(int, m.groups())
    try:
        return datetime.date(y, mo, d).isoformat()
    except ValueError:
        raise ValueError(f"存在しない日付です: {value}") from None


class Resolver:
    """トラッカー名や担当者名などを、Redmine の ID に変換する"""

    def __init__(self, rm, project):
        self.rm = rm
        proj = rm.get(f"/projects/{project}.json?include=trackers,issue_categories")["project"]
        self.project_id = proj["id"]
        self.project_name = proj["name"]
        self.trackers = {t["name"]: t["id"] for t in proj.get("trackers", [])}
        self.categories = {c["name"]: c["id"] for c in proj.get("issue_categories", [])}

        prio = rm.get("/enumerations/issue_priorities.json")["issue_priorities"]
        self.priorities = {p["name"]: p["id"] for p in prio}
        self.statuses = {s["name"]: s["id"] for s in rm.get("/issue_statuses.json")["issue_statuses"]}
        self.versions = {v["name"]: v["id"] for v in rm.get(f"/projects/{project}/versions.json")["versions"]}

        me = rm.get("/users/current.json")["user"]
        self.me_id = me["id"]
        self.me_name = f'{me.get("lastname", "")} {me.get("firstname", "")}'.strip() or me.get("login")
        self.members = {}
        offset = 0
        while True:
            res = rm.get(f"/projects/{project}/memberships.json?limit=100&offset={offset}")
            for m in res["memberships"]:
                who = m.get("user") or m.get("group")
                if who:
                    self.members[who["name"]] = who["id"]
            offset += 100
            if offset >= res.get("total_count", 0):
                break

    @staticmethod
    def _lookup(table, value, label):
        if value in table:
            return table[value]
        choices = "、".join(table) or "(なし)"
        raise ValueError(f"{label}「{value}」が見つかりません。使えるもの: {choices}")

    def tracker(self, v):
        return self._lookup(self.trackers, v, "トラッカー")

    def priority(self, v):
        return self._lookup(self.priorities, v, "優先度")

    def status(self, v):
        return self._lookup(self.statuses, v, "ステータス")

    def version(self, v):
        return self._lookup(self.versions, v, "対象バージョン")

    def category(self, v):
        return self._lookup(self.categories, v, "カテゴリ")

    def assignee(self, v):
        if v in SELF_WORDS:
            return self.me_id, self.me_name
        if v.isdigit():
            return int(v), f"ユーザー#{v}"
        return self._lookup(self.members, v, "担当者(プロジェクトのメンバー)"), v


# ============================================================
# チェックと作成計画
# ============================================================

def build_plan(rows, resolver, rm, use_prefix):
    """ファイルの行を検証し、親から順に並べた作成計画を返す。エラーはまとめて返す"""
    errors = []
    by_key = {}
    for row in rows:
        key = row["key"]
        if not key:
            errors.append(f'{row["_line"]}行目: 行IDが空欄です')
        elif key in by_key:
            errors.append(f'{row["_line"]}行目: 行ID「{key}」が重複しています')
        else:
            by_key[key] = row

    # 既存チケット(#123)を親にする場合の情報
    existing = {}

    for row in by_key.values():
        p = row.get("parent", "")
        if not p:
            continue
        if re.fullmatch(r"#?\d+", p) and p.lstrip("#") not in by_key:
            num = int(p.lstrip("#"))
            if num not in existing:
                try:
                    issue = rm.get(f"/issues/{num}.json")["issue"]
                    existing[num] = issue["subject"]
                except RedmineError as e:
                    errors.append(f'{row["_line"]}行目: 親チケット #{num} を取得できません({e})')
                    existing[num] = None
        elif p not in by_key:
            errors.append(f'{row["_line"]}行目: 親「{p}」がファイルの中に見つかりません')

    # 親から順に並べる(循環のチェックつき)
    order, state = [], {}

    def visit(key, chain):
        if state.get(key) == "done":
            return
        if state.get(key) == "visiting":
            errors.append("親子関係が循環しています: " + " → ".join(chain + [key]))
            return
        state[key] = "visiting"
        p = by_key[key].get("parent", "")
        if p in by_key:
            visit(p, chain + [key])
        state[key] = "done"
        order.append(key)

    for key in by_key:
        visit(key, [])

    # 各行の内容を解決
    plan = []
    full_subject = {}
    for key in order:
        row = by_key[key]
        line = row["_line"]
        item = {"key": key, "line": line, "fields": {"project_id": resolver.project_id}}
        f = item["fields"]
        try:
            f["tracker_id"] = resolver.tracker(row["tracker"])
            item["tracker"] = row["tracker"]
        except ValueError as e:
            errors.append(f"{line}行目: {e}")
            item["tracker"] = row["tracker"]

        # 親と題名
        p = row.get("parent", "")
        parent_subject = None
        if p in by_key:
            item["parent_key"] = p
            parent_subject = full_subject.get(p)
        elif re.fullmatch(r"#?\d+", p):
            num = int(p.lstrip("#"))
            f["parent_issue_id"] = num
            item["parent_existing"] = num
            parent_subject = existing.get(num)

        subject = row["subject"]
        if not subject:
            errors.append(f"{line}行目: 題名が空欄です")
        if use_prefix and parent_subject:
            prefix = parent_subject + SUBJECT_SEPARATOR
            if not subject.startswith(prefix):
                subject = prefix + subject
        if len(subject) > 255:
            errors.append(f"{line}行目: 題名が255文字を超えています")
        f["subject"] = subject
        full_subject[key] = subject

        # 担当者
        try:
            f["assigned_to_id"], item["assignee"] = resolver.assignee(row.get("assignee", ""))
        except ValueError as e:
            errors.append(f"{line}行目: {e}")

        # 日付
        try:
            start = normalize_date(row.get("start_date"))
            due = normalize_date(row.get("due_date"))
            if start:
                f["start_date"] = start
            if due:
                f["due_date"] = due
            if start and due and start > due:
                errors.append(f"{line}行目: 開始日({start})が期日({due})より後になっています")
        except ValueError as e:
            errors.append(f"{line}行目: {e}")

        # 任意の項目
        for col, func, field in [
            ("priority", resolver.priority, "priority_id"),
            ("status", resolver.status, "status_id"),
            ("version", resolver.version, "fixed_version_id"),
            ("category", resolver.category, "category_id"),
        ]:
            if row.get(col):
                try:
                    f[field] = func(row[col])
                    item[col] = row[col]
                except ValueError as e:
                    errors.append(f"{line}行目: {e}")

        if row.get("estimated_hours"):
            try:
                f["estimated_hours"] = float(row["estimated_hours"])
            except ValueError:
                errors.append(f'{line}行目: 予定工数は数値で書いてください: {row["estimated_hours"]}')

        desc = row.get("description", "").replace("\\n", "\n")
        if not desc:
            desc = DESCRIPTION_TEMPLATES.get(row["tracker"], "")
        if desc:
            f["description"] = desc

        plan.append(item)

    return plan, errors


def depth_of(item, plan_by_key):
    """表示用の字下げの深さ(既存チケットの子なら1段深くする)"""
    d = 0
    seen = {item["key"]}
    while item.get("parent_key") in plan_by_key:
        item = plan_by_key[item["parent_key"]]
        if item["key"] in seen:  # 親子関係が循環している(エラーとして別途報告済み)
            break
        seen.add(item["key"])
        d += 1
    if item.get("parent_existing"):
        d += 1
    return d


def print_plan(plan):
    plan_by_key = {i["key"]: i for i in plan}
    for item in plan:
        f = item["fields"]
        indent = "    " * depth_of(item, plan_by_key)
        parent_note = f'(既存 #{item["parent_existing"]} の子)' if item.get("parent_existing") else ""
        details = [f'担当: {item.get("assignee", "-")}']
        if f.get("start_date") or f.get("due_date"):
            details.append(f'{f.get("start_date", "")}〜{f.get("due_date", "")}')
        for col, label in [("priority", "優先度"), ("status", "ステータス"),
                           ("version", "バージョン"), ("category", "カテゴリ")]:
            if item.get(col):
                details.append(f"{label}: {item[col]}")
        print(f'{indent}[{item["key"]}] {item["tracker"]}  {f["subject"]}  {parent_note}')
        print(f'{indent}      ' + " / ".join(details))


# ============================================================
# 作成
# ============================================================

def apply_plan(plan, rm, base_url, log_path):
    created = {}
    log_rows = []
    try:
        for item in plan:
            fields = dict(item["fields"])
            if item.get("parent_key"):
                fields["parent_issue_id"] = created[item["parent_key"]]
            res = rm.post("/issues.json", {"issue": fields})
            issue_id = res["issue"]["id"]
            created[item["key"]] = issue_id
            log_rows.append([item["key"], issue_id, fields["subject"], f"{base_url}/issues/{issue_id}"])
            print(f'  作成: #{issue_id}  [{item["key"]}] {fields["subject"]}')
    except RedmineError as e:
        print(f'\n!! [{item["key"]}](ファイルの{item["line"]}行目)の作成で失敗しました: {e}', file=sys.stderr)
        print("   ここで処理を止めました。これより前のチケットは作成済みです。", file=sys.stderr)
        return log_rows, False
    finally:
        with open(log_path, "w", encoding="utf-8-sig", newline="") as fp:
            w = csv.writer(fp)
            w.writerow(["行ID", "チケット番号", "題名", "URL"])
            w.writerows(log_rows)
    return log_rows, True


# ============================================================
# メイン
# ============================================================

def load_api_key():
    key = os.environ.get("REDMINE_API_KEY", "").strip()
    if not key:
        path = os.path.expanduser("~/.redmine_api_key")
        if os.path.exists(path):
            key = open(path, encoding="utf-8").read().strip()
    if not key:
        raise SystemExit(
            "APIキーが見つかりません。\n"
            "Redmine の「個人設定」→「APIアクセスキー」で表示したキーを、\n"
            "  echo 'キー' > ~/.redmine_api_key && chmod 600 ~/.redmine_api_key\n"
            "のように保存してください。"
        )
    return key


def main():
    ap = argparse.ArgumentParser(description="CSV / Excel から Redmine のチケットを親子付きで一括作成する")
    ap.add_argument("csv", metavar="ファイル", help="チケットを書いた CSV または Excel(.xlsx)ファイル")
    ap.add_argument("--project", "-p", required=True, help="プロジェクトの識別子(例: my-project)")
    ap.add_argument("--sheet", help=f"Excel のシート名(既定: 「{DEFAULT_SHEET}」シート、なければ先頭のシート)")
    ap.add_argument("--apply", action="store_true", help="実際に作成する(付けなければ試し実行)")
    ap.add_argument("--no-prefix", action="store_true", help="題名の頭に親の題名を付けない")
    ap.add_argument("--url", default=DEFAULT_URL, help=f"Redmine の URL(既定: {DEFAULT_URL})")
    args = ap.parse_args()

    rows = read_rows(args.csv, args.sheet)
    if not rows:
        raise SystemExit("ファイルにチケットがありません")

    rm = Redmine(args.url, load_api_key())
    try:
        resolver = Resolver(rm, args.project)
    except RedmineError as e:
        raise SystemExit(f"Redmine の情報を取得できません: {e}")

    plan, errors = build_plan(rows, resolver, rm, use_prefix=not args.no_prefix)

    print(f"プロジェクト: {resolver.project_name}({args.project})")
    print(f"作成予定: {len(plan)}件\n")
    print_plan(plan)

    if errors:
        print(f"\n!! ファイルに {len(errors)} 件の問題があります。修正してからやり直してください。", file=sys.stderr)
        for e in errors:
            print("   - " + e, file=sys.stderr)
        sys.exit(1)

    if not args.apply:
        print("\n(試し実行のため、作成はしていません。作成するには --apply を付けて実行してください)")
        return

    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    log_path = os.path.splitext(args.csv)[0] + f".result-{stamp}.csv"
    print("\n作成を開始します。")
    log_rows, ok = apply_plan(plan, rm, args.url.rstrip("/"), log_path)
    print(f"\n{len(log_rows)}件を作成しました。結果: {log_path}")
    if not ok:
        sys.exit(2)


if __name__ == "__main__":
    main()
