# 岡本パン GitHub Pages テスト版

## フォルダ
- `docs/`：GitHub Pagesで公開する予約画面
- `apps-script/`：既存Apps Scriptへ追加・差し替えするファイル

## 反映手順
1. `apps-script` の4ファイルをApps Scriptプロジェクトへ反映
2. `clasp push`
3. 既存Webアプリを「新しいバージョン」で再デプロイ
4. リポジトリへ `docs` フォルダをpush
5. GitHub Settings → Pages → Deploy from a branch → `test2` / `/docs`
6. 公開URLを確認

既存のApps Script版トップページは残ります。
