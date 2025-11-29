# Slack自動通知シート

Google スプレッドシートのタスク進捗を Slack に自動通知する Apps Script 一式です。Config シートの設定に基づき、即時通知と夜間キューイング／朝送信を中央でまとめて管理します。

## 概要
- Config シートで各拠点シートの URL や通知タイミングを管理し、入力規則でミスを防止。
- onOpen で Config の検証・修復を自動実行、メニューから手動実行も可能。
- トースト＋アラートで開始／完了／エラーをフィードバック。
- 夜間スキャンで期限超過・当日タスクを Queue に蓄積し、朝送信で Slack Webhook へ一括送信。
- Config の `status` / `last_updated` 列に同期結果を記録。

## セットアップ
1. Apps Script プロジェクトに各 `.gs` ファイルをコピーします（`Slack自動通知シート` 配下）。
2. Config シートを開き、メニュー「Central」→「設定シートの検証・修復」を実行してヘッダーと入力規則を適用。
3. Config シートに各拠点の設定を入力します（主な列）:
   - `enabled`: TRUE/FALSE
   - `site_code`: 任意のコード
   - `spreadsheet_url`: 監視対象スプレッドシート
   - `task_sheet` / `members_sheet`: シート名（未指定は `Tasks` / `Members`）
   - `slack_webhook`: 送信先 Incoming Webhook
   - `mode`: `central` / `local` / `none`
   - `notify_timing`: `immediate` / `morning` / `both`
4. Apps Script のトリガーで以下を推奨設定:
   - 時間主導トリガーで `nightlyScanAndQueue`（夜間）
   - 時間主導トリガーで `morningDispatch`（朝）

## 使い方
- メニュー「Central」
  - 「台帳反映（トリガー同期）」: Config から onEdit トリガーを各拠点に作成／削除し、結果を `status` に記録。
  - 「夜間スキャン→Queue」: 手動で夜間スキャンを実行し、Queue に追加。
  - 「朝の送信」: Queue の PENDING を Slack に送信。
  - 「設定シートの検証・修復」: Config の入力規則・ヘッダーを再適用。
- 即時通知: 拠点シートでチェックボックス更新時に `notify_timing` が `immediate` であれば Slack へ即時送信。

## ファイル構成（主なもの）
- `constants.gs`: シート名やヘッダー定義、運用パラメータ。
- `config_repo.gs`: Config 読み書き、入力規則適用、ステータス更新。
- `triggers.gs`: メニュー生成、トリガー同期。
- `queue.gs`: 夜間スキャンで Queue へ積み、朝送信で Slack 送信。
- `handle_edit.gs`: onEdit 即時通知ハンドラ。
- `slack.gs`: Slack への投稿とメッセージ生成。
- `utils.gs`: 共通ユーティリティ（ヘッダー確保、バリデーション、通知など）。

## 運用メモ
- スキャン件数や送信上限は `constants.gs` の `APP_CONFIG` で調整できます。
- キー重複防止で Queue に再登録されない設計です。上限到達時はトーストで警告。
- トリガー同期や送信でエラーがあれば Config の `status` に `ERROR: ...` を記録します。
