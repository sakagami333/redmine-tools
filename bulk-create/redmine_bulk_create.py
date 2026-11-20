#!/usr/bin/env python3
"""
Redmine チケット一括作成ツール(作成・更新)

CSV または Excel(.xlsx)に書いたチケット(親子孫の階層つき)を、
Redmine の REST API で新しく作成する。行IDを #453 のように書いた行は、
既存のチケット #453 の更新として扱う(作成と同じ書式で書ける)。
標準ライブラリだけで動くので、pip でのインストールは不要。

使い方:
  # 試し実行(何も反映しない。内容のチェックと、作成・更新の予定を表示)
  python3 redmine_bulk_create.py tickets.csv --project my-project
  python3 redmine_bulk_create.py tickets.xlsx --project my-project

  # 本番実行(予定を表示し、y で確認してから反映する)
  python3 redmine_bulk_create.py tickets.xlsx --project my-project --apply

  # 確認を省略して反映する
  python3 redmine_bulk_create.py tickets.xlsx --project my-project --apply --yes

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

# 値を「空にする」ときに書く記号(更新で期日や担当者を消すときなど)
CLEAR_WORD = "-"

# 担当者欄でこれらを書くと(新規作成では空欄でも)、自分が担当者になる
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
REQUIRED_COLUMNS = ["key"]  # トラッカー・題名は、新しく作る行でだけ必須(行ごとに確認する)


# ============================================================
# Redmine API
# ============================================================

class RedmineError(Exception):
    def __init__(self, message, status=None):
        super().__init__(message)
        self.status = status  # HTTP のステータスコード(通信できなかった場合は None)


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
            raise RedmineError(f"{method} {path} → HTTP {e.code}: {detail}", status=e.code) from None
        except urllib.error.URLError as e:
            raise RedmineError(f"{self.url} に接続できません: {e.reason}") from None

    def get(self, path):
        return self._request("GET", path)

    def post(self, path, body):
        return self._request("POST", path, body)

    def put(self, path, body):
        return self._request("PUT", path, body)


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
        key = row.get("key", "")
        if key.startswith("#") and not re.fullmatch(r"#\d+", key):
            continue  # 「#」で始まり、数字だけが続くのではない行はメモとして無視(#453 は既存チケットの更新)
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
# チェックと実行計画
# ============================================================

# 既存のチケットの書き方(行ID・親の欄で共通): #453
EXISTING_RE = re.compile(r"#(\d+)")


def _issue_label(issue_id):
    return f"#{issue_id}"


def build_plan(rows, resolver, rm, use_prefix):
    """
    ファイルの行を検証し、親から順に並べた実行計画(作成・更新)を返す。
    エラーはまとめて返す(1件でもあれば、何も実行しない)。
    """
    errors = []

    def err(line, msg):
        errors.append(f"{line}行目: {msg}")

    # --- 行IDの確認と、作成 / 更新の振り分け ---
    by_key = {}
    for row in rows:
        key, line = row["key"], row["_line"]
        if not key:
            err(line, "行IDが空欄です")
            continue
        if key.isdigit():
            err(line, f"行ID「{key}」は数字だけです。既存のチケットを更新するなら「#{key}」、"
                      "新しく作るなら「S1」「T1」のように、数字だけでない名前にしてください")
            continue
        if key in by_key:
            err(line, f"行ID「{key}」が重複しています(1つのチケットは1行にまとめてください)")
            continue
        m = EXISTING_RE.fullmatch(key)
        row["_update"] = int(m.group(1)) if m else None
        by_key[key] = row

    # --- 既存のチケット(更新するもの・親にするもの)を取得 ---
    existing = {}

    def fetch(num, line, purpose):
        if num not in existing:
            try:
                existing[num] = rm.get(f"/issues/{num}.json")["issue"]
            except RedmineError as e:
                existing[num] = None
                if e.status == 404:
                    err(line, f"{purpose} #{num} が見つかりません(番号が違うか、見る権限がありません)")
                else:
                    err(line, f"{purpose} #{num} を取得できません({e})")
        return existing[num]

    for row in by_key.values():
        line = row["_line"]
        if row["_update"]:
            issue = fetch(row["_update"], line, "更新するチケット")
            if issue and issue["project"]["id"] != resolver.project_id:
                err(line, f'#{row["_update"]} は別のプロジェクト(「{issue["project"]["name"]}」)のチケットです。'
                          "--project で指定したプロジェクトのチケットだけ更新できます")
                row["_skip"] = True
        p = row.get("parent", "")
        if not p or p == CLEAR_WORD or p in by_key:
            continue
        m = EXISTING_RE.fullmatch(p)
        if m:
            fetch(int(m.group(1)), line, "親チケット")
        elif p.isdigit():
            err(line, f"親「{p}」:既存のチケットなら「#{p}」と書いてください")
        else:
            err(line, f"親「{p}」がファイルの中に見つかりません")

    # --- 親から順に並べる(循環のチェックつき) ---
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

    # 表示と実行の順番:親のすぐ後に子が続くように並べ直す(親が先になる点は変わらない)
    def tree_parent(key):
        """並べ替え・表示に使う親の行ID(更新で親を変えない行は、今の親がファイルにあればそれ)"""
        row = by_key[key]
        p = row.get("parent", "")
        if p in by_key:
            return p
        if not p and row.get("_update"):
            cur_parent = ((existing.get(row["_update"]) or {}).get("parent") or {}).get("id")
            if cur_parent and f"#{cur_parent}" in by_key:
                return f"#{cur_parent}"
        return None

    children = {}
    for key in order:
        children.setdefault(tree_parent(key), []).append(key)
    file_pos = {k: i for i, k in enumerate(by_key)}
    tree_order, placed = [], set()

    def walk(key):
        if key in placed:
            return
        placed.add(key)
        tree_order.append(key)
        for child in sorted(children.get(key, []), key=file_pos.get):
            walk(child)

    for root in sorted(children.get(None, []), key=file_pos.get):
        walk(root)
    order = tree_order + [k for k in order if k not in placed]  # 循環している行(エラー報告済み)

    # --- 各行の内容を、作成 / 更新の内容に変換 ---
    plan = []
    full_subject = {}  # 行ID → 題名(子の題名の頭に付けるため)

    for key in order:
        row = by_key[key]
        line = row["_line"]
        upd = row["_update"]
        cur = existing.get(upd) if upd else None
        if upd and (cur is None or row.get("_skip")):
            continue  # 取得できなかった(エラーは報告済み)

        item = {
            "key": key, "line": line, "action": "update" if upd else "create",
            "display_parent": tree_parent(key),
            "issue_id": upd, "fields": {}, "changes": [], "summary": [],
            "cur_parent_id": (cur.get("parent") or {}).get("id") if cur else None,
        }
        f = item["fields"]

        def val(col):
            return row.get(col, "")

        def change(label, old, new):
            item["changes"].append((label, old if old not in (None, "") else "(なし)",
                                    new if new not in (None, "") else "(なし)"))

        # トラッカー
        t = val("tracker")
        if t == CLEAR_WORD:
            err(line, "トラッカーは「-」(空にする)にできません")
        elif t:
            try:
                tid = resolver.tracker(t)
                if not upd:
                    f["tracker_id"] = tid
                elif tid != cur["tracker"]["id"]:
                    f["tracker_id"] = tid
                    change("トラッカー", cur["tracker"]["name"], t)
            except ValueError as e:
                err(line, e)
        elif not upd:
            err(line, "トラッカーが空欄です(新しく作るチケットでは必須)")
        item["tracker"] = t or (cur["tracker"]["name"] if cur else "")

        # 親
        p = val("parent")
        parent_subject = None
        if p == CLEAR_WORD:
            if upd and item["cur_parent_id"]:
                f["parent_issue_id"] = ""
                change("親", _issue_label(item["cur_parent_id"]), "")
        elif p in by_key:
            item["parent_key"] = p
            parent_subject = full_subject.get(p)
            if upd:
                pm = EXISTING_RE.fullmatch(p)
                if pm and int(pm.group(1)) != item["cur_parent_id"]:
                    f["parent_issue_id"] = int(pm.group(1))
                    change("親", _issue_label(item["cur_parent_id"]) if item["cur_parent_id"] else "", p)
                elif not pm:
                    change("親", _issue_label(item["cur_parent_id"]) if item["cur_parent_id"] else "",
                           f"[{p}](このファイルで新しく作るもの)")
            elif EXISTING_RE.fullmatch(p):
                item["parent_existing"] = int(EXISTING_RE.fullmatch(p).group(1))
        elif EXISTING_RE.fullmatch(p):
            num = int(EXISTING_RE.fullmatch(p).group(1))
            parent_subject = (existing.get(num) or {}).get("subject")
            item["parent_existing"] = num
            if not upd:
                f["parent_issue_id"] = num
            elif num != item["cur_parent_id"]:
                f["parent_issue_id"] = num
                change("親", _issue_label(item["cur_parent_id"]) if item["cur_parent_id"] else "", p)

        # 題名
        s = val("subject")
        if s == CLEAR_WORD:
            err(line, "題名は「-」(空にする)にできません")
            s = ""
        if upd:
            # 更新では、書いたとおりの題名にする(親の題名は付けない)
            if s and s != cur["subject"]:
                f["subject"] = s
                change("題名", cur["subject"], s)
            full_subject[key] = s or cur["subject"]
        else:
            if not s:
                err(line, "題名が空欄です(新しく作るチケットでは必須)")
            if use_prefix and parent_subject:
                prefix = parent_subject + SUBJECT_SEPARATOR
                if s and not s.startswith(prefix):
                    s = prefix + s
            f["subject"] = s
            full_subject[key] = s
        if len(full_subject[key]) > 255:
            err(line, "題名が255文字を超えています")
        item["subject"] = full_subject[key]

        # 担当者
        a = val("assignee")
        cur_assignee = (cur or {}).get("assigned_to") or {}
        if upd:
            if a == CLEAR_WORD:
                if cur_assignee:
                    f["assigned_to_id"] = ""
                    change("担当者", cur_assignee.get("name"), "")
            elif a:
                try:
                    aid, aname = resolver.assignee(a)
                    if aid != cur_assignee.get("id"):
                        f["assigned_to_id"] = aid
                        change("担当者", cur_assignee.get("name"), aname)
                except ValueError as e:
                    err(line, e)
        else:
            if a == CLEAR_WORD:
                item["summary"].append("担当: (なし)")
            else:
                try:
                    f["assigned_to_id"], aname = resolver.assignee(a)  # 空欄なら自分
                    item["summary"].append(f"担当: {aname}")
                except ValueError as e:
                    err(line, e)

        # 開始日・期日
        dates = {}
        for col, label in [("start_date", "開始日"), ("due_date", "期日")]:
            v = val(col)
            old = (cur or {}).get(col)
            if v == CLEAR_WORD:
                dates[col] = None
                if upd and old:
                    f[col] = ""
                    change(label, old, "")
                continue
            if not v:
                dates[col] = old if upd else None
                continue
            try:
                d = normalize_date(v)
            except ValueError as e:
                err(line, e)
                continue
            dates[col] = d
            if not upd:
                f[col] = d
            elif d != old:
                f[col] = d
                change(label, old, d)
        if dates.get("start_date") and dates.get("due_date") and dates["start_date"] > dates["due_date"]:
            err(line, f'開始日({dates["start_date"]})が期日({dates["due_date"]})より後になります')
        if not upd and (f.get("start_date") or f.get("due_date")):
            item["summary"].append(f'{f.get("start_date", "")}〜{f.get("due_date", "")}')

        # 選択肢の項目(優先度・ステータス・対象バージョン・カテゴリ)
        for col, label, func, field, cur_key, clearable in [
            ("priority", "優先度", resolver.priority, "priority_id", "priority", False),
            ("status", "ステータス", resolver.status, "status_id", "status", False),
            ("version", "対象バージョン", resolver.version, "fixed_version_id", "fixed_version", True),
            ("category", "カテゴリ", resolver.category, "category_id", "category", True),
        ]:
            v = val(col)
            old = (cur or {}).get(cur_key) or {}
            if v == CLEAR_WORD:
                if not clearable:
                    err(line, f"{label}は「-」(空にする)にできません")
                elif upd and old:
                    f[field] = ""
                    change(label, old.get("name"), "")
                continue
            if not v:
                continue
            try:
                vid = func(v)
            except ValueError as e:
                err(line, e)
                continue
            if not upd:
                f[field] = vid
                item["summary"].append(f"{label}: {v}")
            elif vid != old.get("id"):
                f[field] = vid
                change(label, old.get("name"), v)

        # 予定工数
        v = val("estimated_hours")
        old = (cur or {}).get("estimated_hours")
        if v == CLEAR_WORD:
            if upd and old is not None:
                f["estimated_hours"] = ""
                change("予定工数", old, "")
        elif v:
            try:
                hours = float(v)
                if not upd:
                    f["estimated_hours"] = hours
                elif old is None or abs(float(old) - hours) > 1e-9:
                    f["estimated_hours"] = hours
                    change("予定工数", old, hours)
            except ValueError:
                err(line, f"予定工数は数値で書いてください: {v}")

        # 説明
        v = val("description").replace("\\n", "\n")
        old = (cur or {}).get("description") or ""
        if upd:
            if v == CLEAR_WORD:
                if old:
                    f["description"] = ""
                    change("説明", "(あり)", "")
            elif v and v != old:
                f["description"] = v
                change("説明", "(あり)" if old else "", f"書き換え({len(v)}文字)")
        else:
            if v == CLEAR_WORD:
                v = ""
            elif not v:
                v = DESCRIPTION_TEMPLATES.get(item["tracker"], "")
            if v:
                f["description"] = v

        if not upd:
            f["project_id"] = resolver.project_id
        plan.append(item)

    def line_of(msg):
        m = re.match(r"(\d+)行目", msg)
        return int(m.group(1)) if m else 0

    errors.sort(key=line_of)  # 行の順に並べる
    return plan, errors


def depth_of(item, plan_by_key):
    """表示用の字下げの深さ(ファイルにない既存チケットの子なら1段深くする)"""
    d = 0
    seen = {item["key"]}
    while (item.get("parent_key") or item.get("display_parent")) in plan_by_key:
        item = plan_by_key[item.get("parent_key") or item.get("display_parent")]
        if item["key"] in seen:  # 親子関係が循環している(エラーとして別途報告済み)
            break
        seen.add(item["key"])
        d += 1
    if item.get("parent_existing"):
        d += 1
    return d


def is_noop(item):
    """更新する行で、変える項目が何もないか"""
    return item["action"] == "update" and not item["changes"]


def print_plan(plan):
    plan_by_key = {i["key"]: i for i in plan}
    for item in plan:
        indent = "    " * depth_of(item, plan_by_key)
        pad = indent + "      "
        parent_note = ""
        if item.get("parent_existing") and item.get("parent_key") not in plan_by_key:
            parent_note = f'(既存 #{item["parent_existing"]} の子)'
        if item["action"] == "create":
            print(f'{indent}[{item["key"]}] 新規  {item["tracker"]}  {item["subject"]}  {parent_note}')
            print(pad + " / ".join(item["summary"]))
        else:
            print(f'{indent}[{item["key"]}] 更新  {item["tracker"]}  {item["subject"]}')
            if is_noop(item):
                print(pad + "(変更なし)")
            for label, old, new in item["changes"]:
                print(f"{pad}{label}: {old} → {new}")


# ============================================================
# 実行(作成・更新)
# ============================================================

def _verify_update(issue, fields):
    """更新後のチケットを確かめ、指定どおりに変わらなかった項目を返す"""
    def ref_id(name):
        return (issue.get(name) or {}).get("id")

    actual = {
        "tracker_id": ref_id("tracker"),
        "status_id": ref_id("status"),
        "priority_id": ref_id("priority"),
        "assigned_to_id": ref_id("assigned_to"),
        "fixed_version_id": ref_id("fixed_version"),
        "category_id": ref_id("category"),
        "parent_issue_id": ref_id("parent"),
        "subject": issue.get("subject"),
        "start_date": issue.get("start_date"),
        "due_date": issue.get("due_date"),
        "estimated_hours": issue.get("estimated_hours"),
        "description": issue.get("description") or "",
    }
    labels = {
        "tracker_id": "トラッカー", "status_id": "ステータス", "priority_id": "優先度",
        "assigned_to_id": "担当者", "fixed_version_id": "対象バージョン", "category_id": "カテゴリ",
        "parent_issue_id": "親", "subject": "題名", "start_date": "開始日", "due_date": "期日",
        "estimated_hours": "予定工数", "description": "説明",
    }
    mismatched = []
    for field, wanted in fields.items():
        got = actual.get(field)
        if wanted == "":
            ok = got in (None, "")
        elif field == "estimated_hours":
            ok = got is not None and abs(float(got) - float(wanted)) < 1e-9
        elif field == "description":
            ok = got.replace("\r\n", "\n") == wanted.replace("\r\n", "\n")
        else:
            ok = got == wanted
        if not ok:
            mismatched.append(labels.get(field, field))
    return mismatched


def apply_plan(plan, rm, base_url, log_path):
    ids = {}          # 行ID → チケット番号
    log_rows = []
    warn_count = 0
    item = None
    try:
        for item in plan:
            if item["action"] == "create":
                fields = dict(item["fields"])
                if item.get("parent_key"):
                    fields["parent_issue_id"] = ids[item["parent_key"]]
                res = rm.post("/issues.json", {"issue": fields})
                issue_id = res["issue"]["id"]
                ids[item["key"]] = issue_id
                log_rows.append([item["key"], "作成", issue_id, fields["subject"], f"{base_url}/issues/{issue_id}", ""])
                print(f'  作成: #{issue_id}  [{item["key"]}] {fields["subject"]}')
                continue

            # 更新
            issue_id = item["issue_id"]
            ids[item["key"]] = issue_id
            fields = dict(item["fields"])
            pk = item.get("parent_key")
            if pk and not EXISTING_RE.fullmatch(pk) and ids[pk] != item["cur_parent_id"]:
                fields["parent_issue_id"] = ids[pk]  # このファイルで作った親に付け替える
            if not fields:
                log_rows.append([item["key"], "変更なし", issue_id, item["subject"], f"{base_url}/issues/{issue_id}", ""])
                continue
            rm.put(f"/issues/{issue_id}.json", {"issue": fields})
            after = rm.get(f"/issues/{issue_id}.json")["issue"]
            mismatched = _verify_update(after, fields)
            note = ""
            if mismatched:
                warn_count += 1
                note = "反映されなかった項目: " + "、".join(mismatched)
                print(f'  更新: #{issue_id}  [{item["key"]}] {after["subject"]}')
                print(f"    !! {note}(ワークフローの制限や、子チケットから自動計算される項目の可能性があります)")
            else:
                print(f'  更新: #{issue_id}  [{item["key"]}] {after["subject"]}')
            log_rows.append([item["key"], "更新", issue_id, after["subject"], f"{base_url}/issues/{issue_id}", note])
    except RedmineError as e:
        what = "作成" if item["action"] == "create" else "更新"
        print(f'\n!! [{item["key"]}](ファイルの{item["line"]}行目)の{what}で失敗しました: {e}', file=sys.stderr)
        print("   ここで処理を止めました。これより前の行は反映済みです。", file=sys.stderr)
        return log_rows, False, warn_count
    finally:
        with open(log_path, "w", encoding="utf-8-sig", newline="") as fp:
            w = csv.writer(fp)
            w.writerow(["行ID", "処理", "チケット番号", "題名", "URL", "注意"])
            w.writerows(log_rows)
    return log_rows, True, warn_count


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
    ap = argparse.ArgumentParser(description="CSV / Excel から Redmine のチケットを親子付きで一括作成・更新する")
    ap.add_argument("csv", metavar="ファイル", help="チケットを書いた CSV または Excel(.xlsx)ファイル")
    ap.add_argument("--project", "-p", required=True, help="プロジェクトの識別子(例: my-project)")
    ap.add_argument("--sheet", help=f"Excel のシート名(既定: 「{DEFAULT_SHEET}」シート、なければ先頭のシート)")
    ap.add_argument("--apply", action="store_true",
                    help="実際に反映する(予定を表示し、確認してから作成・更新する。付けなければ試し実行)")
    ap.add_argument("--yes", "-y", action="store_true", help="--apply のとき、作成前の確認を省略する")
    ap.add_argument("--no-prefix", action="store_true", help="題名の頭に親の題名を付けない")
    ap.add_argument("--url", default=DEFAULT_URL, help=f"Redmine の URL(既定: {DEFAULT_URL})")
    args = ap.parse_args()
    base_url = args.url.rstrip("/")

    rows = read_rows(args.csv, args.sheet)
    if not rows:
        raise SystemExit("ファイルにチケットがありません")

    rm = Redmine(base_url, load_api_key())
    try:
        resolver = Resolver(rm, args.project)
    except RedmineError as e:
        if e.status == 404:
            raise SystemExit(
                f"プロジェクト「{args.project}」が {base_url} に見つかりません。\n"
                "プロジェクトの識別子(URL の /projects/〇〇 の〇〇)と、接続先(--url)を確認してください。"
            )
        if e.status in (401, 403):
            raise SystemExit(
                f"{base_url} に接続を断られました(HTTP {e.status})。\n"
                "APIキーが正しいか、「管理 → 設定 → API」で REST が有効になっているか、"
                "そのプロジェクトを見る権限があるかを確認してください。"
            )
        raise SystemExit(f"Redmine の情報を取得できません: {e}")

    plan, errors = build_plan(rows, resolver, rm, use_prefix=not args.no_prefix)

    n_create = sum(1 for i in plan if i["action"] == "create")
    n_update = sum(1 for i in plan if i["action"] == "update" and not is_noop(i))
    n_noop = sum(1 for i in plan if is_noop(i))

    print(f"接続先    : {base_url}")
    print(f"プロジェクト: {resolver.project_name}({args.project})")
    counts = f"新規作成 {n_create}件 / 更新 {n_update}件"
    if n_noop:
        counts += f" / 変更なし {n_noop}件"
    print(f"予定      : {counts}\n")
    print_plan(plan)

    if errors:
        print(f"\n!! ファイルに {len(errors)} 件の問題があります。修正してからやり直してください。", file=sys.stderr)
        for e in errors:
            print("   - " + e, file=sys.stderr)
        sys.exit(1)

    if n_create == 0 and n_update == 0:
        print("\n作成・更新するものはありません。")
        return

    if not args.apply:
        print("\n(試し実行のため、何も反映していません。反映するには --apply を付けて実行してください)")
        return

    # 本番実行:上の一覧を確認してもらってから反映する
    if not args.yes:
        what = "・".join(x for x in [f"新規作成 {n_create}件" if n_create else "",
                                     f"更新 {n_update}件" if n_update else ""] if x)
        question = f"\n上の内容({what})を、{base_url} の「{resolver.project_name}」に反映します。よろしいですか? [y/N]: "
        try:
            answer = input(question).strip().lower()
        except EOFError:
            answer = ""
        if answer not in ("y", "yes"):
            print("中止しました(何も反映していません)。")
            return

    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    log_path = os.path.splitext(args.csv)[0] + f".result-{stamp}.csv"
    n = 2
    while os.path.exists(log_path):  # 同じ秒に実行した結果を上書きしない
        log_path = os.path.splitext(args.csv)[0] + f".result-{stamp}-{n}.csv"
        n += 1
    print("\n反映を開始します。")
    log_rows, ok, warn_count = apply_plan(plan, rm, base_url, log_path)
    done_create = sum(1 for r in log_rows if r[1] == "作成")
    done_update = sum(1 for r in log_rows if r[1] == "更新")
    print(f"\n新規作成 {done_create}件、更新 {done_update}件を反映しました。結果: {log_path}")
    if warn_count:
        print(f"!! {warn_count}件で、指定どおりに変わらなかった項目があります。上の「!!」の行を確認してください。")
    if not ok:
        sys.exit(2)


if __name__ == "__main__":
    main()
