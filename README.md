# 登壇諾否マイページ 共通管理版 v3.9

## v3.9のメール仕様
- 送信元SMTPアカウントはシステム全体で1つだけ設定します。
- 学会ごとにSMTPアカウントやパスワードを設定する必要はありません。
- 回答時は、回答者を To、学会事務局を CC、Reply-To も学会事務局にして同一メールを送信します。
- 学会ごとに設定するメール項目は「運営事務局メールアドレス」「件名」「回答内容より前の案内文」です。
- 差出人表示名は「＜学会名＞ 運営事務局」として自動生成されます。
- 回答内容、備考、初回登録情報、回答日時、マイページURLはシステムが自動生成します。差込タグは不要です。
- 未回答者へのリマインド送信は行いません。先生別URL・未回答者一覧をExcel出力し、既存のVBA送信で利用します。

## Streamlit Secrets（共通で1回のみ）
```toml
ADMIN_PASSWORD = "管理画面用パスワード"
TOKEN_SECRET = "十分に長いランダム文字列"

SMTP_HOST = "smtp.gakkai.co.jp"
SMTP_PORT = 587
SMTP_USERNAME = "quo@gakkai.co.jp"
SMTP_PASSWORD = "送信用アカウントのパスワード"
SMTP_SECURITY = "starttls"
SMTP_FROM_EMAIL = "quo@gakkai.co.jp"
```

※ SMTPパスワードはGitHubへ書かず、Streamlit Community Cloud の Settings → Secrets にのみ保存してください。
