#!/usr/bin/env python3
"""
Redmine チケット一括作成ツール

CSV に書いたチケット(親子孫の階層つき)を、Redmine の REST API で作成する。
標準ライブラリだけで動くので、pip でのインストールは不要。

使い方:
  # 試し実行(作成はしない。内容のチェックと、作成予定の一覧を表示)
  python3 redmine_bulk_create.py tickets.csv --project my-project

  # 本番実行(実際に作成する)
  python3 redmine_bulk_create.py tickets.csv --project my-project --apply

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
# CSV の読み込み
# ============================================================

def read_csv(path):
    """UTF-8(BOM付き含む)と Shift_JIS(Excel の標準)の両方に対応"""
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp932"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise SystemExit("CSV の文字コードを判別できません(UTF-8 か Shift_JIS で保存してください)")

    reader = csv.DictReader(text.splitlines())
    header_map = {}
    for name in reader.fieldnames or []:
        clean = name.strip()
        for key, aliases in COLUMN_ALIASES.items():
            if clean in aliases:
                header_map[name] = key
    missing = [k for k in REQUIRED_COLUMNS if k not in header_map.values()]
    if missing:
        names = ", ".join(COLUMN_ALIASES[k][0] for k in missing)
        raise SystemExit(f"CSV に必須の列がありません: {names}")

    rows = []
    for line_no, rec in enumerate(reader, start=2):
        row = {key: (rec.get(src) or "").strip() for src, key in header_map.items()}
        if not any(row.values()):
            continue  # 空行は無視
        if row.get("key", "").startswith("#") and not row.get("tracker"):
            continue  # 「#」で始まるコメント行は無視
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
    """CSV の行を検証し、親から順に並べた作成計画を返す。エラーはまとめて返す"""
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
            errors.append(f'{row["_line"]}行目: 親「{p}」がCSVの中に見つかりません')

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
        print(f'\n!! [{item["key"]}](CSV {item["line"]}行目)の作成で失敗しました: {e}', file=sys.stderr)
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
    ap = argparse.ArgumentParser(description="CSV から Redmine のチケットを親子付きで一括作成する")
    ap.add_argument("csv", help="チケットを書いた CSV ファイル")
    ap.add_argument("--project", "-p", required=True, help="プロジェクトの識別子(例: my-project)")
    ap.add_argument("--apply", action="store_true", help="実際に作成する(付けなければ試し実行)")
    ap.add_argument("--no-prefix", action="store_true", help="題名の頭に親の題名を付けない")
    ap.add_argument("--url", default=DEFAULT_URL, help=f"Redmine の URL(既定: {DEFAULT_URL})")
    args = ap.parse_args()

    rows = read_csv(args.csv)
    if not rows:
        raise SystemExit("CSV にチケットがありません")

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
        print(f"\n!! CSV に {len(errors)} 件の問題があります。修正してからやり直してください。", file=sys.stderr)
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
