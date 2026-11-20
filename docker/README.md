# Redmine を WSL2 + Docker で動かす構成

Windows の WSL2(Ubuntu)に Docker Engine を入れ、Redmine 7.0 と PostgreSQL 16 をコンテナで動かすための構成です。Docker Desktop は使いません(企業規模によっては有償ライセンスが必要なため)。

| ファイル | 内容 |
|---|---|
| [docker-compose.yml](docker-compose.yml) | Redmine と PostgreSQL のコンテナの定義 |
| [Dockerfile](Dockerfile) | 公式イメージに、プラグインの gem のビルドに必要なツールを追加 |
| [.env.example](.env.example) | パスワードと鍵の雛形(`.env` にコピーして使う) |
| [backup.sh](backup.sh) | 日次バックアップ(DB と添付ファイル、14日分を保持) |

## 構成

```
Windows
└─ WSL2(Ubuntu)
    └─ Docker Engine
        ├─ Redmine コンテナ     http://localhost:3000(この PC からのみ)
        └─ PostgreSQL コンテナ
```

## 置き方

```
~/redmine-local/
├── docker-compose.yml
├── Dockerfile
├── .env
├── backup.sh
├── config/configuration.yml
├── data/files/
└── plugins/
```

データは WSL2 のホーム配下に置きます。Windows 側(`/mnt/c/…`)に置くと、読み書きが大幅に遅くなります。

## 移行元から持ってくるもの

```bash
# 移行元のサーバで
sudo -u postgres pg_dump -Fc redmine > redmine.dump
cd /var/lib/redmine && tar czf ~/redmine-files.tar.gz files && tar czf ~/redmine-plugins.tar.gz plugins
cp config/configuration.yml ~/
```

## 起動

```bash
cd ~/redmine-local
docker compose up -d db
docker compose exec -T db pg_restore -U redmine -d redmine --no-owner --no-privileges < redmine.dump
docker compose up -d --build
docker compose logs -f redmine      # "Listening on" が出たら起動完了
```

## つまずいた点と対処

| 症状 | 原因と対処 |
|---|---|
| Redmine が再起動を繰り返す(`chown: … Read-only file system`) | `configuration.yml` を読み取り専用(`:ro`)で渡していた。公式イメージは起動時に所有者を変えるため、`:ro` を外す |
| `docker pull` が `network is unreachable` / `no such host` で失敗する | WSL2 の DNS がときどき応答に失敗し、Docker が IPv6 だけで繋ごうとしていた。数回やり直すと通る。Docker に OS 標準の名前解決を使わせる設定も入れた(下記) |
| `docker` が `permission denied` | ユーザーが `docker` グループに入っていない。`sudo usermod -aG docker $USER` のあと、PowerShell で `wsl --terminate <ディストリビューション名>` して入り直す |
| Docker の再起動後、コンテナが自動で戻らない | `containerd` と `docker` を同時に再起動したときに起きた。Docker を再起動したら `docker compose ps` で確認し、止まっていれば `docker compose start` |
| WSL が数十秒で止まり、Redmine も止まる | ターミナルを閉じると WSL が自動停止する。Windows 側の `%USERPROFILE%\.wslconfig` に下記を書く(BOM なしで保存) |

Docker に OS 標準の名前解決を使わせる設定:

```bash
sudo mkdir -p /etc/systemd/system/docker.service.d /etc/systemd/system/containerd.service.d
printf '[Service]\nEnvironment="GODEBUG=netdns=cgo"\n' | sudo tee /etc/systemd/system/docker.service.d/dns-cgo.conf /etc/systemd/system/containerd.service.d/dns-cgo.conf
sudo systemctl daemon-reload
sudo systemctl restart containerd docker
```

WSL の自動停止を止める設定(`%USERPROFILE%\.wslconfig`):

```ini
[general]
instanceIdleTimeout=-1
```

## PC 起動時に自動で立ち上げる

コンテナは `restart: unless-stopped` なので、WSL が起動すれば Redmine も上がります。PowerShell で、ログオン時に WSL を起動するタスクを登録します(`Ubuntu-26.04` は自分のディストリビューション名に置き換え)。

```powershell
$action   = New-ScheduledTaskAction -Execute "wsl.exe" -Argument "-d Ubuntu-26.04 -u root -e true"
$trigger  = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
Register-ScheduledTask -TaskName "Start WSL Redmine" -Action $action -Trigger $trigger -Settings $settings
```

## バックアップ

```bash
chmod +x backup.sh
REDMINE_BACKUP_COPY_DIR=/mnt/c/Users/<ユーザー名>/RedmineBackup ./backup.sh
```

cron で毎日実行する例は、`backup.sh` の冒頭のコメントを参照してください。バックアップを PC の外(外付けストレージなど)にも、ときどきコピーしておくと安心です。
