# 登壇諾否マイページ 共通管理版 v4.1

## 今回の更新
- 外部PostgreSQLによる永続保存に対応しました。
- Streamlit Community Cloudの再起動・再デプロイ後も、学会設定・先生情報・依頼・回答が残せます。
- 招聘状等の指定様式アップロードもDB内へ保存します（10MBまで）。
- 管理画面「運用設定」からシステム全体バックアップ（JSON）をダウンロード／復元できます。
- DATABASE_URL が未設定なら従来どおりローカルSQLiteで動作します。

## 推奨構成
本番では、外部PostgreSQL（例：Supabase）を1つだけ用意して全学会で共用します。
学会ごとにDBを作る必要はありません。

Streamlitの Settings → Secrets に以下を追加します。

```toml
DATABASE_URL = "postgresql://ユーザー名:パスワード@ホスト名:5432/postgres"
```

Supabaseを使う場合は、プロジェクトの Connect 画面に表示される **Session pooler** の接続文字列をそのまま使ってください。
接続文字列内の `[YOUR-PASSWORD]` は、Supabaseプロジェクト作成時のDBパスワードに置き換えます。

## 永続DBへ切り替える前の手順
1. 旧版の管理画面 → 運用設定 → 「システムバックアップをダウンロード」
2. v4.1へ更新
3. 外部PostgreSQLを用意し、Secretsへ `DATABASE_URL` を追加
4. アプリを再起動
5. 管理画面 → 運用設定 → 「バックアップから復元」

これで既存の学会設定・先生情報・依頼・回答を新しい永続DBへ移せます。

## 共通Gmail SMTP設定
回答受付メールは共通Gmailから送信し、各学会の事務局メールアドレスをCC／Reply-Toに使います。

```toml
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USERNAME = "共通送信用Gmailアドレス"
SMTP_PASSWORD = "Googleの16桁アプリパスワード"
SMTP_SECURITY = "starttls"
SMTP_FROM_EMAIL = "共通送信用Gmailアドレス"
SMTP_TLS_VERIFY = true
```

メール送信仕様：
- To：回答者本人
- CC：学会ごとの運営事務局メールアドレス
- Reply-To：学会ごとの運営事務局メールアドレス
- From：共通Gmail
