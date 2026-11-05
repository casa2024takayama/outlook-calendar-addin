# outlook-calendar-addin（メールから予定）

Outlook（Mac の新しい Outlook / Web / Windows）のアドイン。開いているメールの本文から日付を読み取り、確認パネルで直してから「新しい予定」画面を入力済みで開く。**保存は Outlook でユーザーが押す**（予定を直接書き込まない）。

- 権限は `ReadItem`（開いているメールを読むだけ）。Microsoft Graph は使わない
- 本文はパネルの中だけで処理し、外部には送らない
- 日付の読み取りルールは 社内記事用の予定ブックマークレット v0.5.0 を移植（`src/parser.js`）

## 使い方

1. メールを開き、ツールバー（または「…」→アプリ）の **メールから予定** を押す → 右にパネルが出る
2. 候補の件名・日時・場所を確認して直す（「＋ 手動で追加」もあり）
3. **予定を作成** → 入力済みの予定画面が開く → **保存**
4. パネルのピン留めをすると、別のメールを選ぶたびに読み直す

### 終日の予定

Outlook アプリの予定画面には「終日」を直接指定できない（`displayNewAppointmentForm` の仕様）。そのため 0:00〜翌日0:00 で開くので、画面で **終日** をオンにして保存する。
代わりに「Web版で終日として開く」を押すと、ブラウザの Outlook Web で終日の予定画面が開く。

## 読み取りルール（元のブックマークレットからの変更点）

- 返信・転送の引用（`-----Original Message-----`、`From:`／`差出人:`、`On … wrote:`、`>` 行）から下は読まない
- 年の基準はメールの**受信日**。受信日より2か月以上前の日付は翌年とみなす
- 件名の日付（【11/20締切】など）は、本文に同じ日が出ていなければ追加
- 件名に付けるメール名は `RE:`／`FW:` と日付入りの括弧を外したもの
- 場所（行頭「場所」「会場」「開催場所」）は、締切系の予定には付けない
- 過去の日付・期間の中の1日・年の誤記らしきものは、注記を付けて表示（過去は下にまとめる）
- URL の中の日付（`…/2026/11/05/…`）は読まない。「Ver 1/2」「参加率は 3/4」のような分数（年・曜日・時刻なしの「数/数」で、直前が版・率・割合・点数など）は日付にしない（v0.1.2）

## 登録済みの記録

「予定を作成」を押した候補は登録済みとして記録し、同じメールを開くと「✓ 登録済み」と表示する。記録先はアドインの localStorage（キー `olCalAddin.v1`、メールの internetMessageId ごと）。Outlook で実際に保存したかは確認していない。

## 開発

```bash
node build.mjs https://<配置先>     # src/ と assets/ → docs/、manifest.xml に配置先URLを入れる
node test/run.mjs                    # 自動テスト（下の「自動テスト」）
```

- ブラウザで `docs/taskpane.html` を直接開くと（Outlook の外）、件名と本文を貼り付けて試せる「試し画面」になる
- プレビュー：`projects/.claude/launch.json` の `outlook-addin`（http://localhost:8812/taskpane.html）
- アドイン ID：`8ba43ede-d65b-4c29-99b8-7ee60d015534`。直したら `build.mjs` の `VERSION` を上げる

## 自動テスト

```bash
node test/run.mjs              # 単体テスト＋画面テスト（ヘッドレス Chrome）
node test/run.mjs --manifest   # ＋ manifest 検証（Microsoft の検証サービスに送る）
node test/run.mjs create-timed # シナリオを指定して一部だけ
```

- 単体テスト：`test/*.test.mjs`（node だけで動く）
- 画面テスト：`test/harness/`。`src/taskpane.html` から本物の Office.js を外し、偽物（`office-mock.js`）を差し込んだページを `test/.out/` に作って、Chrome を画面なしで動かす。Outlook に渡した値（予定画面・送信内容）を記録して確かめる
  - シナリオは `test/harness/scenarios.js`。1シナリオ＝1回のページ読み込みで、状態は毎回まっさら
  - 表示されているかは hidden 属性ではなく CSS の結果で判定する（Outlook の中で試し入力欄が見えていた不具合を見逃さないため）
- 実際の Outlook・Slack には何も送らない。Outlook 実機での見た目と、本物の Office.js との違いは確かめられない
- 直したら、push する前に `node test/run.mjs --manifest` を通す（manifest のバージョンが 1.0 未満だとインストールで弾かれる）

## 導入

Outlook →「アプリを取得（アドインを取得）」→「個人用アドイン」→「カスタムアドインを追加」→「ファイルから追加」で `docs/manifest.xml` を選ぶ。

配置先は GitHub Pages（`main` ブランチの `docs/`）：https://casa2024takayama.github.io/outlook-calendar-addin/taskpane.html
Outlook は manifest に書いたこの URL からパネルを毎回読み込む。直したら `node build.mjs https://casa2024takayama.github.io/outlook-calendar-addin` → コミット → push で反映（manifest の入れ直しは、ボタン名やアイコンなど manifest を変えたときだけ）。

公開リポジトリなので、コードやテストに社内の名前・実際のメール文面を入れないこと。
