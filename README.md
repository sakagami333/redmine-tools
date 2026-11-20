# redmine-tools

Redmine を便利に使うための小さな道具をまとめたリポジトリです。

| フォルダ | 内容 |
|---|---|
| [bulk-create](bulk-create/) | CSV / Excel から、親子孫の階層つきでチケットを一括作成するツール(Python 3、REST API 使用) |
| [view-customize](view-customize/) | View customize プラグインに登録して使う画面の改良(デザイン調整、ヘッダーのリンク、外部リンクのボタンなど) |
| [tampermonkey](tampermonkey/) | 新規チケット作成時に、トラッカー・題名・説明などを自動入力するブラウザ用スクリプト |
| [docker](docker/) | Redmine を WSL2 + Docker で動かす構成と、バックアップのスクリプト |

各ツールの使い方は、それぞれのフォルダの README.md を参照してください。

## 動作確認環境

- Redmine 7.0
- Python 3(標準ライブラリのみ。追加のインストールは不要)
