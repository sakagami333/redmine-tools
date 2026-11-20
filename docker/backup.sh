#!/bin/bash
# Redmine の日次バックアップ(DB と添付ファイル)
#
# docker-compose.yml と同じフォルダに置いて実行する。
# 添付ファイルはコンテナの中で固めるので、sudo は不要。
#
# 環境変数 REDMINE_BACKUP_COPY_DIR を設定すると、そのフォルダにもコピーする
# (例: Windows 側の /mnt/c/Users/<ユーザー名>/RedmineBackup)
#
# cron の例(毎日12時):
#   0 12 * * * /home/<ユーザー名>/redmine-local/backup.sh >> /home/<ユーザー名>/redmine-local/backup/backup.log 2>&1
set -euo pipefail
cd "$(dirname "$0")"

DATE=$(date +%F)
KEEP_DAYS=14
COPY_DIR="${REDMINE_BACKUP_COPY_DIR:-}"
mkdir -p backup

/usr/bin/docker compose exec -T db pg_dump -U redmine -Fc redmine > "backup/redmine-$DATE.dump"
/usr/bin/docker compose exec -T redmine tar czf - -C /usr/src/redmine files > "backup/files-$DATE.tar.gz"

DIRS=(backup)
if [ -n "$COPY_DIR" ]; then
  mkdir -p "$COPY_DIR"
  cp "backup/redmine-$DATE.dump" "backup/files-$DATE.tar.gz" "$COPY_DIR/"
  DIRS+=("$COPY_DIR")
fi

find "${DIRS[@]}" -maxdepth 1 -name "redmine-*.dump" -mtime +"$KEEP_DAYS" -delete
find "${DIRS[@]}" -maxdepth 1 -name "files-*.tar.gz" -mtime +"$KEEP_DAYS" -delete
echo "$(date '+%F %T') backup OK"
