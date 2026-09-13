# 登壇諾否マイページ 共通管理版 v4.0

## SMTP互換対応
Shurikenの既存設定に合わせて、587 + STARTTLS + SMTP AUTH に対応しています。

今回の設定例：

```toml
SMTP_HOST = "211.13.204.15"
SMTP_PORT = 587
SMTP_USERNAME = "送信元アカウントのユーザー名"
SMTP_PASSWORD = "送信元アカウントのパスワード"
SMTP_SECURITY = "starttls"
SMTP_FROM_EMAIL = "送信元メールアドレス"
SMTP_TLS_VERIFY = false
```

`SMTP_TLS_VERIFY=false` は、IPアドレス接続時に証明書のホスト名が一致しない既存SMTPサーバー向けの互換設定です。通常のSMTPでは true を推奨します。

メール送信仕様：
- To: 回答者本人
- CC: 学会ごとの事務局メールアドレス
- Reply-To: 学会ごとの事務局メールアドレス
- From: 共通SMTP送信元
