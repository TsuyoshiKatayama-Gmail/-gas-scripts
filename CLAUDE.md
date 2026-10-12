# CLAUDE.md

このファイルは、本プロジェクトで Claude Code が作業する際のガイドラインです。

## プロジェクト概要

- Google スプレッドシートと Google Apps Script（GAS）を利用したアプリケーションを開発します。
- スプレッドシートをデータベース兼UIとして扱い、GAS でロジック・自動化・外部連携を実装します。

## 技術スタック

- **言語**: Google Apps Script（JavaScript ベース）
- **データ保存先**: Google スプレッドシート
- **実行環境**: Google Workspace（GAS ランタイム V8）
- **主要サービス**: `SpreadsheetApp`、`PropertiesService`、`UrlFetchApp`、トリガー（時間主導型・イベント主導型）など

## コーディング規約

- 関数・変数名は英語のキャメルケース（例: `getActiveSheet`、`updateRow`）を使用します。
- コメントは日本語で記述し、処理の意図がわかるようにします。
- スプレッドシートのシート名・列名などの固定値は、ファイル先頭の定数（`const`）としてまとめて定義します。
- スプレッドシートへのアクセスはループ内で繰り返さず、`getValues()` / `setValues()` で一括読み書きして実行時間を短縮します。
- API キーやトークンなどの機密情報はコードに直書きせず、`PropertiesService`（スクリプトプロパティ）で管理します。
- `.gs` ファイルは機能単位で分割します（例: `main.gs`、`sheetService.gs`、`utils.gs`）。

## ディレクトリ構成（想定）

```
gas-scripts/
├── CLAUDE.md          # 本ファイル
├── .clasp.json        # clasp 設定（ローカル⇔GAS 同期を行う場合）
├── appsscript.json    # GAS マニフェスト
└── src/               # .gs / .html ソース
```

> ローカルと GAS プロジェクトの同期には [clasp](https://github.com/google/clasp) の利用を推奨します。

## Git 運用ルール（重要）

**コードを変更するたびに、必ず GitHub へプッシュします。**

1. 作業前に最新を取得する: `git pull`
2. コードを変更したら、変更内容がわかる単位でコミットする:
   ```bash
   git add .
   git commit -m "変更内容を簡潔に記述"
   ```
3. **コミット後は必ず GitHub へプッシュする**:
   ```bash
   git push
   ```
4. 1回の変更（1つの機能追加・修正）ごとにコミット＆プッシュを行い、変更を GitHub 上に残す。
5. コミットメッセージは日本語で、「何を・なぜ変更したか」が伝わるように書く。

> 現時点ではまだ Git リポジトリが初期化されていません。最初に以下を実行してください。
> ```bash
> git init
> git remote add origin <GitHubリポジトリのURL>
> ```

## Claude Code への指示

- 回答・コミットメッセージ・コメントはすべて日本語で記述する。
- コードを変更した後は、上記の Git 運用ルールに従い、コミットとプッシュまで実施する。
- スプレッドシートの構造（シート名・列構成）を変更する場合は、影響範囲を確認してから実施する。
