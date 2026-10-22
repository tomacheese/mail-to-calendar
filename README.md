# mail-to-calendar

Exchange Online の受信トレイに届いた予約・予定関連のメールを Microsoft Foundry で解析し、Outlook のメインカレンダーへ予定として自動登録する。Azure Functions (v4, Node.js, Flex Consumption) 上で 5 分ごとの Timer Trigger として動作する。

- メール取得: Microsoft Graph の Inbox Delta（新着のみ。初回は基準点を作るだけで過去メールは登録しない）
- 解析: 既存の Microsoft Foundry（API キー認証）で分類・日時抽出し、アプリ側で日時・根拠・予約の同一性を再検証
- 登録: 確定は `busy`、仮予約は `【仮】` + `tentative`。変更は更新、取消は削除（削除済みアイテムに移動）
- 予約は「施設名 + 予約番号」で同一性を判定する。別の施設の予約番号と一致しても、既存の予定の更新・削除は行わず保留する（同じ施設でも表記が大きく異なる場合は別の施設として扱う）
- 手動編集された予定は上書きしない。管理対象は自動作成した予定のみ

## セットアップ

### 必要環境

- Node.js（バージョンは `.node-version` 参照）
- pnpm（`preinstall` で `only-allow pnpm` を強制）

```bash
pnpm install
```

### 環境変数

ローカルは `.env`（Git 管理外）、Functions では `local.settings.json` の `Values` / アプリ設定に設定する。雛形は [`.env.example`](./.env.example)。

| 変数名                                                        | 必須 | 説明                                                                 |
| ------------------------------------------------------------- | ---- | -------------------------------------------------------------------- |
| `GRAPH_TENANT_ID` / `GRAPH_CLIENT_ID` / `GRAPH_CLIENT_SECRET` | ✅   | Graph アプリ登録（Client ID は GUID、Secret は値）                   |
| `GRAPH_MAILBOX`                                               | ✅   | 対象メールボックスの UPN                                             |
| `FOUNDRY_ENDPOINT` / `FOUNDRY_DEPLOYMENT` / `FOUNDRY_API_KEY` | ✅   | 既存 Foundry の接続情報                                              |
| `STATE_STORAGE_CONNECTION`                                    | ✅   | 状態保存先（ローカルは Azurite、本番はアプリのストレージ）           |
| `CALENDAR_WRITE_MODE`                                         | -    | `dry-run`（既定）/ `create-only` / `full`                            |
| `MAIL_POLLING_ENABLED`                                        | -    | Timer による自動監視（既定 `false`）                                 |
| `PROCESS_EXISTING_MESSAGES`                                   | -    | 初回に既存メールを処理するか（既定 `false`）                         |
| `DEFAULT_EVENT_DURATION_MINUTES`                              | -    | 終了時刻が不明な場合の所要時間（既定 60）                            |
| `FOUNDRY_MAX_CALLS_PER_RUN`                                   | -    | 1 回の実行での推論呼び出し上限（アプリ既定 50、Bicep の本番既定 20） |
| `STATE_TABLE_PREFIX`                                          | -    | テーブル名の接頭辞（同一ストレージ内の分離用）                       |

### 開発コマンド

```bash
pnpm start check        # Graph / Foundry の疎通確認（実メール本文は読まない）
pnpm start scan         # 受信トレイの差分を処理（モードは CALENDAR_WRITE_MODE）
pnpm start sample 8     # 既存メールを dry-run で解析
pnpm start report       # dry-run で記録した計画を表示
pnpm start:func         # ビルドしてローカルの Functions ホストを起動
pnpm build              # tsc ビルド（dist/functions/*.js）
pnpm test               # jest（Azurite を自動起動）
pnpm test:e2e           # 合成メールで実 Foundry を呼ぶ E2E（Graph はスタブ、課金あり）
pnpm lint               # prettier / eslint / tsc
pnpm fix                # prettier / eslint の自動修正
```

## ディレクトリ構成

```text
src/
  main.ts                # ローカル CLI（check / analyze / sample / report / scan）
  config.ts              # 環境変数の検証
  functions/timer.ts     # 5 分間隔の Timer Trigger（共通エンジンを呼ぶ薄い入口）
  core/                  # 予約処理エンジン・検証・状態遷移
  integrations/          # Graph / Foundry / Azure Tables
test/                    # jest（Azurite は global-setup で自動起動）
infra/                   # Bicep（main.bicep / main.bicepparam / README.md）
```

## デプロイ

インフラ構築手順は [`infra/README.md`](./infra/README.md)、アプリのデプロイは [`.github/workflows/azure-functions-deploy.yml`](./.github/workflows/azure-functions-deploy.yml)（`main` / `master` への push）を参照。本番の初期モードは必ず `dry-run` とし、メインカレンダーへの書き込みは別途承認のうえで切り替える。

## License

[MIT](./LICENSE)
