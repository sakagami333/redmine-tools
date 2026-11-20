# 動作確認したバージョンに固定したい場合は redmine:7.0.3 のように指定する
FROM redmine:7.0

# プラグインの gem(redmine_checklists が使う redmineup など)を
# 起動時の bundle install でビルドできるよう、ビルドツールを追加する
RUN set -eux; \
    apt-get update; \
    apt-get install -y --no-install-recommends build-essential libpq-dev; \
    rm -rf /var/lib/apt/lists/*
