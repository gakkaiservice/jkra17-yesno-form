# 登壇諾否マイページ 共通管理版 v4.1.1

v4.1 の PostgreSQL/Supabase 永続保存版を、外部DB向けに速度改善した修正版です。

## v4.1.1 の修正
- Excel取込を1件ずつDB往復する方式から一括登録に変更
- 先生別URL画面のプロフィール取得を1人ずつ行わず一括取得に変更
- Excel取込中はスピナーを表示し、完了件数を画面に残す
- 既存の機能・DB構造・Secrets（DATABASE_URL等）は変更なし

GitHubには `app.py` / `requirements.txt` / `README.md` を上書きしてください。
