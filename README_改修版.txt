【同封ファイル】
1. kaikei_uchiwake_app_pro.py ・・・ 改修済みアプリ本体
2. requirements_pro.txt ・・・ 依存関係
3. initial_balance_template.xlsx ・・・ 初期残高テンプレート
4. journal_master_template.xlsx ・・・ 全期間仕訳マスタテンプレート
5. demo_initial_balance.xlsx / demo_journal_master.xlsx ・・・ デモExcel
6. demo_initial_balance_utf8sig.csv / demo_journal_master_utf8sig.csv ・・・ UTF-8 BOM付きCSV
7. demo_initial_balance_cp932.csv / demo_journal_master_cp932.csv ・・・ Excel向けCP932 CSV

【今回の改修ポイント】
- 初期残高＋全期間仕訳マスタ対応
- 対象月選択
- 期首残高 / 当月増加 / 当月減少 / 期末残高の自動計算
- 相手先別内訳
- 期首算出用の過去履歴と当月仕訳明細のドリルダウン
- Excel出力
- デザインをカード・タブ中心の見やすいUIへ変更

【推奨運用】
- 初回のみ初期残高を準備
- 以後は全期間仕訳マスタに毎月追記
- アプリで対象月を選んで確認

【Streamlit Cloudに置く場合】
- メインファイル: kaikei_uchiwake_app_pro.py
- requirements_pro.txt は必要に応じて requirements.txt にリネームして配置

【注意】
- 相手先は表記ゆれ厳禁（例: A社 / 株式会社A / A(株) を混在させない）
- 仕訳IDは重複防止のためユニーク推奨
