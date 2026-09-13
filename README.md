# 登壇諾否マイページ 共通管理版 v3.0

複数の学会を1つのStreamlitアプリで管理するための共通版です。

## できること

- 管理画面から新しい学会を追加
- 学会ごとに初回質問項目をON/OFF
  - ふりがな
  - 会員区分
  - 緊急連絡先
  - 氏名・所属等の修正依頼
  - 招聘状／派遣依頼状セット（JKRA17だけON等が可能）
- Excelを取り込み、先生ごとの専用URLを自動生成
- 1人に複数依頼がある場合は同じマイページにまとめて表示
- 既回答は固定表示、追加依頼だけ回答可能
- 初回回答後は、2回目以降「諾否＋備考」だけ入力可能
- 回答一覧・先生情報をExcel出力

## ローカル起動

```powershell
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

## Streamlit Community Cloud の Secrets（本番公開時）

Streamlitの App settings → Secrets に次を設定してください。

```toml
ADMIN_PASSWORD = "十分に長い管理者パスワード"
TOKEN_SECRET = "ランダムな長い文字列"
YESNO_BASE_URL = "https://あなたのアプリ名.streamlit.app"
```

`TOKEN_SECRET` は一度設定したら変更しないでください。変更すると先生専用URLのキーが変わります。

## 新しい学会を作る手順

1. `?admin=1` で管理画面へログイン
2. 「＋ 新しい学会」
3. 学会名・会期・会場・初回質問項目を設定
4. 「学会管理」→「Excel取込」で指定演題／登壇者Excelをアップロード
5. 「先生別URL」でURL一覧を確認

GitHubリポジトリ作成やStreamlitの新規Deployは、学会ごとには不要です。

## 招聘状について

「招聘状／派遣依頼状セットを使う」をONにした学会だけ表示されます。
承諾が1件以上ある初回回答時にだけ質問し、2回目以降は登録済み内容を表示のみとします。

## 重要：v3.0の保存先

v3.0は共通化の動作確認版で、回答データの保存先はSQLiteです。
Streamlit Community Cloudでは、再起動・再デプロイ時にSQLiteやアップロードファイルが失われる可能性があります。
**本番の回答受付を開始する前に、次版でGoogle Sheets等の永続保存先へ切り替えてください。**
