# 登壇諾否マイページ 共通管理版 v4.2

## v4.2 の変更点

- 永続保存先を **Googleスプレッドシート** に変更しました。
- Streamlit の再起動・再デプロイ後も、スプレッドシートからデータを自動復元します。
- 学会設定、先生情報、依頼、回答、アップロードされた招聘状様式を保存します。
- 学会ごとにGoogle連携を設定する必要はありません。**システム全体で最初の1回だけ**です。
- 管理画面「運用設定」から、現在データの手動保存／読み込みもできます。
- JSONバックアップ機能も残しています。

## Googleスプレッドシート側に自動作成されるシート

- `conferences`：学会設定
- `people`：先生マスター・初回登録情報
- `requests`：依頼マスター
- `responses`：新システムでの回答
- `stored_files`：アップロードファイル（分割保存）
- `_system`：同期日時など

## 最初の1回だけ必要なGoogle設定

1. Googleスプレッドシートを1つ新規作成します。
   - 例：`登壇諾否マイページ_データ` 
2. Google Cloud でサービスアカウントを1つ作成し、JSONキーを発行します。
3. 作成したスプレッドシートを、サービスアカウントの `client_email` に **編集者** として共有します。
4. Streamlit Community Cloud の `Settings → Secrets` に以下を追加します。

```toml
GOOGLE_SHEETS_SPREADSHEET_ID = "スプレッドシートID"
GOOGLE_SHEETS_AUTO_SYNC = true

[gcp_service_account]
type = "service_account"
project_id = "JSON内のproject_id"
private_key_id = "JSON内のprivate_key_id"
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "JSON内のclient_email"
client_id = "JSON内のclient_id"
token_uri = "https://oauth2.googleapis.com/token"
```

※ JSONキーのファイル自体をGitHubへアップロードしないでください。

## 既存データを初回だけスプシへ移す

Google連携後、管理画面の `運用設定` を開き、

**「現在のデータをスプシへ保存」**

を1回押してください。

その後は、学会作成・Excel取込・回答登録・先生情報更新・学会削除などの変更時に自動同期します。

## メール送信

Gmail SMTPなどの共通送信元を1つ設定します。

```toml
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USERNAME = "共通送信用Gmailアドレス"
SMTP_PASSWORD = "Googleの16桁アプリパスワード"
SMTP_SECURITY = "starttls"
SMTP_FROM_EMAIL = "共通送信用Gmailアドレス"
SMTP_TLS_VERIFY = true
```

回答時は、

- To：回答した先生
- CC：各学会の運営事務局
- Reply-To：各学会の運営事務局

で自動送信します。

## 注意

Google Sheets APIの通信エラーが起きた場合、管理画面の「運用設定」から手動同期を再実行できます。
アップロードファイルはGoogle Sheetsのセル上限を避けるため、Base64化した内容を複数行に分割して保存します。
