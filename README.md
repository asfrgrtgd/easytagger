# Local LM Tag Generator
これはobsidianの拡張です

## 最初のセットアップ
1. `src/main.ts` を開く
2. 次の2つを自分の環境に合わせて変更する
   - `LMSTUDIO_BASE_URL`（例: `http://192.168.10.105:1234`）
   - `LMSTUDIO_MODEL`（例: `gpt-oss-swallow-20b-rl-v0.1`）
3. このフォルダで `npm run build`

## 使い方
1. タグ付けしたいノートを開く
2. 左の `Tag Generator` パネルを開く（またはコマンド `Open Tag generator sidebar`）
3. `New tag mode` を選ぶ
   - `なし`: 既存タグだけ使う
   - `承認式`: 新規候補を確認して選んで適用
   - `自動追加`: 新規候補を自動で適用
4. `Run Tag generation` を押す
5. （必要なら）`Auto append new tags to allowed list` を ON にして、新規採用タグを `templates/allowed-tags.md` に追記する
