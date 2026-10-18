# Exchange Online 予約メール → Outlook 予定表自動登録

**要件定義・基本設計・PoC／実装計画 v0.2（完全版）**  
更新日：2026-10-17  
ステータス：**設計中／PoC未着手／本番変更なし**  
対象：Exchange Online + Microsoft Graph + Azure Functions Flex Consumption + 別テナントの既存Microsoft Foundry

> **文書の位置づけ**  
> 2026-10-17までの会話で確定した要件、現在採用を推奨する設計、未確定事項、Azure実環境調査結果を統合したもの。  
> 「確定」はユーザーの明示した要件、「推奨」はPoCで検証しながら変更可能な初期設計を意味する。  
> 認証情報・メール本文などの秘密値は本書に含めない。**本書はデプロイの承認ではない。**

## 0. v0.1からの変更点

1. **Microsoft Graphの認証を、ローカルPoC・本番の共通性を重視して「Entra IDアプリ登録 + Client Secret」に変更（現行推奨方針）**。Graph自体に汎用APIキー方式はない。SecretでOAuth 2.0アクセストークンを取得する。
2. **Microsoft Foundryは、別テナント・別サブスクリプションの既存リソースを再利用**。APIキー認証。新規Foundryリソースは作成しない。
3. **Azure Key Vaultは使わない**。ローカルは `.env` / `local.settings.json`（Git管理外）、本番はAzure Functionsの**アプリ設定（環境変数）**で秘密情報を管理する。
4. **ローカルPoCと本番で同じ認識・照合ロジック、Graphアプリ権限、Foundry API仕様を使用**する。Timer Triggerは共通エンジンの実行入口にとどめる。
5. **Azuriteと本番Storageの状態を分離**し、特にGraph Deltaカーソルと予定IDを共有しない。必要なら同じAzure Tables SDKを使い、接続先のみ切り替える。
6. `dry-run`（標準）→ `create-only` → `full` の**書き込み制御モード**を明記。テストでメインカレンダーを誤変更しない。
7. **既存Foundryのモデル／デプロイ名を確認してから採用モデルを決定**。`model-router`は候補であって確定ではない。
8. 環境変数、実装責務、PoC手順、認証／権限設定、受け入れ条件、未解決事項を追記・改訂。

## 1. 目的と非目的

### 1.1 目的

Exchange Onlineの**受信トレイ（Inbox）**に届く予約関連メールを取得し、Microsoft Foundryで予約状態・日時・施設・予約詳細を判定して、**既存のOutlookメインカレンダー**へ自動登録する。予約変更・キャンセルにも段階的に対応する。

対象は病院に限定せず、歯科・美容院・飲食店・宿泊・交通・イベント・その他の明確な予約とする。**医療関連の予約詳細も予定本文に含める**。

### 1.2 非目的／明確な除外

- Gmailへの転送・Gmailの予約抽出機能への依存。
- 全メールフォルダー監視。**受信トレイ以外に即時振り分けされるメールは原則対象外**。
- 初回起動時に、過去に届いたメールを一括して予定登録すること。
- 不明な開始日時をAIが推測して確定予定として登録すること。
- AIが管理外の既存手動予定を無断で更新・削除すること。
- 初期MVPでのPDF・画像添付のOCR／読解の完全対応。
- Key Vault導入やテナント間IDフェデレーション。
- 今回の設計文書作成をもって本番への変更を許可すること。

## 2. 決定事項・前提条件

| 項目 | 内容 | ステータス |
|---|---|---|
| メール基盤 | Exchange Online（既存メールアドレスを維持） | **確定** |
| 監視範囲 | Inboxのみ。ゴミ箱は監視しない | **確定** |
| 処理対象 | 病院を含む幅広い予約メール | **確定** |
| 登録先 | 既存Outlookメインカレンダー | **確定** |
| 不確実な予約 | 予約の可能性が高く、開始日時が一意なら仮予定として登録 | **確定要件の具体化・推奨** |
| 終了時刻 | 明示情報・妥当な根拠から決められない場合、開始＋60分 | **確定** |
| 医療情報 | 予約関連の詳細を予定本文へ含める | **確定** |
| 成功通知 | 登録完了通知を出さない | **確定** |
| 初回処理 | 新着のみ。過去メールはPoC用の明示的サンプルに限る | **確定** |
| 開発方針 | PoCと実装を並行し、段階的に自動書き込みを有効化 | **確定** |
| Functions | 専用Azure Functions Flex Consumption、Main Azure subscription | **推奨・採用方針** |
| リージョン | Japan East、Always Readyは有効化しない | **推奨** |
| 実行間隔 | 5分ごとのTimer Trigger | **初期案** |
| Graph認証 | Entra IDアプリ登録 + OAuth Client Credentials + Client Secret | **現行推奨（ローカル／本番共通）** |
| Graph権限 | Exchange Online Application RBACで対象メールボックスを限定 | **推奨／権限設定未実施** |
| Foundry | **別Entraテナント・別サブスクリプションの既存Foundry** | **確定** |
| Foundry認証 | 既存FoundryのAPIキー | **確定** |
| Foundryの国外推論 | Globalデプロイを含め許容 | **確定** |
| 秘密値保存 | ローカルはGit管理外の環境変数、本番はFunctionsアプリ設定。Key Vaultなし | **確定** |
| Foundryモデル | 既存デプロイの性能とAPI互換性を確認して決める | **未確定** |
| 初回書き込み | `dry-run`を標準、メインカレンダーへの書き込みは別途明示承認 | **推奨／安全要件** |

## 3. システム構成と責務

```text
             Exchange Online（メインEntraテナント）
                            │
               Microsoft Graph API / OAuth
                 Client ID + Client Secret
                            │
                ┌───────────▼────────────┐
ローカルPoC ────►│ 共通予約処理エンジン       │◄──── Azure Functions本番
CLI / テスト     │ TypeScript（推奨）       │      Timer Trigger / 5分
                │ ・Inbox Delta取得        │
                │ ・メール前処理／重複排除  │
                │ ・予約分類・構造化        │
                │ ・時刻／根拠検証          │
                │ ・予約の同一性判定        │
                │ ・安全な状態遷移          │
                └──────┬──────────┬──────┘
                       │          │
        HTTPS + API Key │          │ 状態保存
                       ▼          ▼
           既存Microsoft Foundry   Azure Tables互換Storage
           （別テナント）          ローカル：Azurite
           モデルの新規作成は不要   本番：Azure Storage
                       │          │
                       └─────┬────┘
                             │
                 Microsoft Graph Calendar API
                   CALENDAR_WRITE_MODEで制御
                             │
                     既存メインカレンダー
```

**配置と料金の分離：**

- **Main Azure subscription：** 専用Functions、Storage、必要な監視、Microsoft Graph接続。
- **Foundryの別サブスクリプション／別テナント：** 既存Foundryと推論料金。
- 別テナントでも**HTTPS + APIキー**ならFunctions側のManaged Identity連携は不要。Foundry側でキー認証が有効かつネットワーク接続が許可されていることが前提。
- 独立した専用Function Appを使用。既存のMCP／ウォッチ用アプリには混在させない。

## 4. 開発と本番の環境同一性／分離

| 項目 | ローカルPoC（Pine等） | Azure本番 |
|---|---|---|
| コアロジック | 同じTypeScriptライブラリ | 同じビルド成果物 |
| 起動 | CLI／単体・結合テスト | Timer Trigger |
| Graph | 同じEntraアプリIDとアプリ権限を使用 | 同じEntraアプリIDとアプリ権限を使用 |
| Graph Secret | ローカル環境変数。必要なら開発用Secretを別発行 | Functionsアプリ設定。開発用とは別Secretを推奨 |
| Foundry | 既存エンドポイント、同じAPI仕様 | 同じエンドポイント、同じAPI仕様 |
| Foundry Key | ローカル環境変数（Git管理外） | Functionsアプリ設定 |
| Storage | **Azurite／テスト専用の状態データ** | **実Azure Storage／本番状態データ** |
| Deltaカーソル | PoC専用（本番と共有しない） | 本番専用 |
| 予定書き込み | **`dry-run`が既定** | 初期`dry-run`。承認後に変更 |
| 自動ポーリング | 原則無効。CLIで1回ずつ実行 | Timerで5分ごと |
| ログ | 秘密値・原文非出力 | 秘密値・原文非出力、保持期間を制限 |

ローカルPoCに本番のメール読み取りを許す一方、**勝手にメインカレンダーを書き換えない**。同じコードと認証方式を採用しつつ、書き込み権限の使用をモードで明示的に制御する。

### 4.1 実行モード（推奨）

| `CALENDAR_WRITE_MODE` | Graphでのメール取得 | Foundry解析 | 予定の操作 |
|---|---|---|---|
| `dry-run` | 可 | 可 | **書き込まない**。作成・更新・取消案のみ出力 |
| `create-only` | 可 | 可 | 自動作成の新規予定のみ作る（既存更新なし） |
| `full` | 可 | 可 | 対象予約の新規作成・安全条件付き更新・取消 |

- **フェイルクローズ：** 未設定・不正なモードでは`dry-run`または起動拒否とし、`full`へ暗黙に移行しない。
- `full`にしても**システム管理外の既存予定は変更不可**。
- CLIの書き込みテストは、テストデータ・識別可能な件名・明示承認を条件とする。
- 本番有効化フラグ、処理対象メールボックス、環境名を起動ログに表示。ただし秘密値は表示しない。

### 4.2 タイマーから処理を分離する

実装例の責務分割：

```text
src/
  functions/timer.ts          # 5分間隔で共通エンジンを呼ぶ薄い入口
  cli/scan.ts                 # ローカルで単発実行・範囲指定・dry-run
  core/process-mail.ts        # 予約処理の共通エンジン
  core/booking-state.ts       # 予約同一性・状態遷移
  core/validation.ts          # 時刻・構造化出力・安全条件
  integrations/graph.ts       # トークン取得・Graph Mail / Calendar
  integrations/foundry.ts     # APIキーによる推論呼び出し
  integrations/state.ts       # Azure Tables互換の状態保存
  config.ts                   # 環境変数検証（Zod等）
tests/
  fixtures/                   # 匿名化・合成メール
  unit/
  integration/
infra/
  main.bicep                  # 専用Functions／Storage等
```

TypeScriptは**提案**であり、実装方式の最終承認前である。

## 5. 認証とアクセス権限

### 5.1 Graph：Client Secret方式（現在の推奨）

**GraphにFoundryのような固定APIキーはない。** Entra IDのアプリ登録を作成し、OAuth 2.0 Client CredentialsフローでGraphアクセストークンを取得する。

- `GRAPH_TENANT_ID`：Exchange Online所属のメインテナントID。
- `GRAPH_CLIENT_ID`：専用アプリ登録のApplication (client) ID。
- `GRAPH_CLIENT_SECRET`：同アプリ登録のClient Secret。**秘密情報**。
- OAuthトークンの要求先はメインテナントのトークンエンドポイント、Graphの対象スコープは`https://graph.microsoft.com/.default`。
- Microsoft Graphアクセス先は`/users/{mailbox-id}/...`等の**アプリ認証対応エンドポイント**。サインインユーザーを要求する`/me`は使用しない。
- ローカルPoC・本番で同じアプリ登録・ロールを使用できる。**Secretだけ開発用と本番用で別発行すると、漏えい時の切り分けが容易**。この場合も権限と動作は共通。
- Secretには有効期限がある。更新期日を記録し、期限前に再発行→環境変数差し替え→疎通確認→旧Secret破棄する。

**Exchange Online Application RBAC（推奨）：**

1. Entra IDアプリ登録と、そのService Principalを確認。
2. Exchange Online側で必要に応じてService Principalへの参照を登録。
3. 対象メールボックスだけを含むManagement Scope（または適切な管理単位）を設定。
4. `Application Mail.Read`と`Application Calendars.ReadWrite`のロール割り当てを行う。
5. `Test-ServicePrincipalAuthorization`で対象内／対象外を検証し、**実際のGraphリクエストでも両方を確認**。
6. **同じ権限についてEntra ID側にテナント全体への無制限なApplication Permissionを付与しない**。Application RBACは無制限なEntra側権限を上書きしない。

> この方式はアプリ登録・Secret・Exchange Application RBACの組み合わせ。Entra ID側で広域Application Permission＋管理者同意を与える方式とは**意図的に区別**し、実際のGraph権限判定をPoCで確認する。Application RBACの設定にはExchange側の適切な管理／委任権限が必要。

### 5.2 Foundry：別テナントの既存リソースにAPIキーで接続（確定）

- 既存Foundryの推論エンドポイントURL、デプロイ名、API形式を確認。
- `FOUNDRY_API_KEY`をHTTPSリクエストの認証に使用。Azure OpenAI形式のREST APIでは通常`api-key`ヘッダー。
- **Foundry側のEntraテナントID・クライアントシークレットは不要**（APIキー認証を使うため）。
- モデルが未デプロイなら、まず既存リソース内で利用可能なデプロイを確認。新しいモデルのデプロイは必要な場合のみ、別途承認を得て実施。
- エンドポイント形式（Azure OpenAI互換 v1／デプロイ指定型／Foundry推論エンドポイント）を**実物で確認**し、SDK・パス・バージョンを決める。
- Globalデプロイによる国外処理はユーザー許容済み。ただしモデルの推論費用は**Foundry側サブスクリプション**に発生。

### 5.3 Storage認証（推奨）

- PoCはAzuriteの`UseDevelopmentStorage=true`等を使用。
- 本番はAzure Storage。実装上は環境変数で接続情報を切り替える。
- 本番StorageはManaged Identity + Azure RBACを第一候補とするが、Functionsホスト・デプロイ用Storageには実際のプラン要件を満たす接続設定が必要。**Graphと同じClient Secret方式に無理に統一しない**。
- 目的はAPI認証方式の形式的な統一でなく、**同じStorage SDK／データモデル／テスト**での動作再現性。

### 5.4 Secretの保管（Key Vaultなし：確定）

- ローカル：`.env`や`local.settings.json`を使用し、Git除外・ファイル権限制限を行う。`.env.example`には**変数名とダミー値のみ**。
- 本番：Azure Functionsの**Application settings（環境変数）**にSecret値を保存。Key Vault参照は使用しない。
- 設定値・秘密値をCLIの標準出力、GitHub Actionsログ、エラー本文、Application Insightsのカスタム属性へ出さない。
- アプリ設定の閲覧・更新権限を必要最小限にする。Secret漏えい時の無効化・再発行手順を準備する。
- Secretをこの文書・チャット・Issue・プルリクエストに貼り付けない。実値の投入方法は、ユーザー自身による管理画面設定など**安全な経路を事前合意**する。

## 6. 環境変数一覧

### 6.1 アプリが参照する変数

| 変数名 | 用途 | 秘密値 | 必須性 |
|---|---|---|---|
| `APP_ENV` | `local`／`production`等 | いいえ | 必須 |
| `GRAPH_TENANT_ID` | Exchange Online側のEntra IDテナント | いいえ | 必須 |
| `GRAPH_CLIENT_ID` | Graphアプリ登録のClient ID | いいえ | 必須 |
| `GRAPH_CLIENT_SECRET` | OAuth Client Secret | **はい** | 必須 |
| `GRAPH_MAILBOX` | 対象メールボックスIDまたはUPN | 個人情報の可能性 | 必須 |
| `FOUNDRY_ENDPOINT` | 別テナントFoundry推論エンドポイント | いいえ | 必須 |
| `FOUNDRY_DEPLOYMENT` | 既存モデルデプロイ名 | いいえ | 必須 |
| `FOUNDRY_API_KEY` | Foundry APIキー | **はい** | 必須 |
| `FOUNDRY_API_VERSION` | 旧／バージョン指定型APIの場合のバージョン | いいえ | APIにより必須 |
| `STATE_STORAGE_CONNECTION` | アプリ状態用Storageの接続設定 | 内容次第 | 必須 |
| `CALENDAR_WRITE_MODE` | `dry-run`／`create-only`／`full` | いいえ | 必須 |
| `MAIL_POLLING_ENABLED` | ローカルの自動監視を抑止 | いいえ | 推奨 |
| `PROCESS_EXISTING_MESSAGES` | 初回の過去メール予定化防止 | いいえ | 推奨（既定`false`） |
| `DEFAULT_EVENT_DURATION_MINUTES` | 所要時間不明時の60分 | いいえ | 推奨 |
| `MAIL_POLL_INTERVAL` | スケジュール等 | いいえ | 任意 |
| `FOUNDRY_MAX_CALLS_PER_RUN` | 1回の実行で許容する推論呼び出し上限 | いいえ | 任意 |

※`AzureWebJobsStorage`、Functions WorkerやDeployment関連の**プラットフォーム必須設定**は、この表とは別にBicep等で設定する。

### 6.2 ローカル`.env.example`（秘密値を含まない）

```dotenv
APP_ENV=local
GRAPH_TENANT_ID=YOUR_MAIN_TENANT_ID
GRAPH_CLIENT_ID=YOUR_APP_CLIENT_ID
GRAPH_CLIENT_SECRET=SET_LOCALLY_DO_NOT_COMMIT
GRAPH_MAILBOX=YOUR_MAILBOX_UPN

FOUNDRY_ENDPOINT=https://your-resource.openai.azure.com/
FOUNDRY_DEPLOYMENT=YOUR_EXISTING_DEPLOYMENT
FOUNDRY_API_KEY=SET_LOCALLY_DO_NOT_COMMIT
# FOUNDRY_API_VERSION=...

STATE_STORAGE_CONNECTION=UseDevelopmentStorage=true
CALENDAR_WRITE_MODE=dry-run
MAIL_POLLING_ENABLED=false
PROCESS_EXISTING_MESSAGES=false
DEFAULT_EVENT_DURATION_MINUTES=60
```

本番では同じ変数名をFunctionsアプリ設定へ投入し、Storageを本番接続先へ切り替える。**`.env`の本番用コピーはGitに置かない。**

## 7. メール取得・初期同期・信頼性

- **監視対象はInboxのみ**。受信トレイ＋ゴミ箱の実測件数は容量見積もりに利用するが、ゴミ箱は監視しない。
- `GET /users/{id}/mailFolders/inbox/messages/delta`等の**フォルダー単位のGraph Delta**を使用。
- 初回のDeltaラウンドは現在のInbox内容を返し得る。**初回処理時は既存メールを予定化せず、最終`@odata.deltaLink`を基準点として確立する**。
- 初期化中に新しいメールが到着した際の境界条件に備え、初回同期開始時刻・受信時刻・処理済み記録を利用し、取りこぼし／二重登録をPoCで検証する。
- 続く同期では`@odata.nextLink`を最後までたどり、`@odata.deltaLink`を保存する。既読変更・移動・削除などのイベントが混在するため「差分＝新着メール」とはみなさない。
- 変更があったメールの本文取得・AI処理前に、必要な識別子を取得し、**処理完了または再処理可能な永続状態ができてからDeltaカーソルを確定**する。
- Graph 429（`Retry-After`）、一時的な5xx、ネットワーク障害、Foundry遅延に指数バックオフ等で対応する。
- GraphメッセージのID変動を考慮し、必要に応じて`Prefer: IdType="ImmutableId"`を使用。`internetMessageId`も補助キーとする。
- **ローカルPoCと本番はDeltaカーソル・処理済みメール表を共有しない**。ローカルで対象メールを検査しても本番の進捗は変えない。
- 受信後5分以内に別フォルダーに移動したメールを見逃すことがある。これはInboxのみの監視というユーザー方針に伴う既知の制約。
- ローカル実行時は自動ポーリングを既定OFFにし、CLIの明示実行で検証する。

## 8. AIの解析仕様

### 8.1 分類（PoCで検証）

- `ignore`：単なる案内・日時の列挙・予約と無関係。
- `confirmed`：予約の成立がメールから確認できる。
- `tentative`：仮予約・申込受付・成立が不確かだが、予約の可能性が高く、日時が一意。
- `update`：既存予約の変更通知。
- `cancel`：既存予約の取消通知。
- `needs_review`：予約らしいが日時や対象が一意でない／判断不能。

### 8.2 登録の判定表

| 受信メールの意味 | 日時の一意性 | カレンダー操作 |
|---|---|---|
| 確定予約 | 一意 | **確定予定**（`showAs: busy`） |
| 仮予約・申込受付 | 一意 | **仮予定**（`showAs: tentative`、件名`【仮】`） |
| 予約の可能性が高いが確信がない | 一意 | **仮予定** |
| 候補日時の提示のみ | 複数または未確定 | 登録しない |
| 開始日時不明 | 不明 | 登録しない・保留 |
| 複数の候補時刻が同等 | 特定できない | 登録しない・保留 |
| 無関係メール | — | 無視 |
| 変更メール | 元予約を一意に照合できる | 管理対象予定を更新 |
| 取消メール | 元予約を一意に照合できる | 安全な取消表示（方式は未確定） |

「曖昧なら仮予定」は**予約成立の確度**に適用し、**存在しない日時の推測は許可しない**。

### 8.3 AI出力スキーマの主要フィールド（推奨）

```json
{
  "action": "confirmed | tentative | update | cancel | ignore | needs_review",
  "bookings": [
    {
      "provider": "施設・店舗名",
      "subject": "カレンダー件名候補",
      "start_at": "ISO 8601 or null",
      "end_at": "ISO 8601 or null",
      "timezone": "Asia/Tokyo",
      "location": "住所・会場・URL等",
      "reservation_number": "予約番号 or null",
      "details": "予約内容・受付方法・持ち物等",
      "evidence": ["原文の短い根拠の列挙"],
      "target_reservation_hint": "変更・取消の照合材料"
    }
  ],
  "reason": "短い判定理由"
}
```

上記は**概念図**であり、有効なJSON Schemaの完成形ではない。実装時には枚数、型、日付表現、null許容、各actionの必須フィールドを確定し、Structured Outputsに対応するモデルなら利用する。アプリ側で**必ず**再バリデーションする。

### 8.4 日時・時間帯

- 元メールの明示日時と文脈に基づいて正規化。国内予約の原則はJST、海外予約は現地タイムゾーンを尊重。
- 開始時刻が一意でなければ予定を作らない。
- 明示された終了時刻・所要時間は優先。推定不能なら**開始＋60分**。
- 相対日付（「来週火曜」等）は受信日時と文脈で一意に決められる場合のみ採用。
- 1通に複数の予約がある場合は**予約単位**で処理。
- HTML本文のノイズを整理し、引用履歴、署名、メール本文中の疑似命令を信頼しない。

## 9. カレンダー予定の作成・変更・取消

### 9.1 新規予定

- 書き込み先は**既存メインカレンダー**。
- 予約名、施設、日時、場所、予約番号、受付・持ち物、注意事項などの**予約詳細を本文に構造化して保存**。医療予約も含む。
- ただし、メール全原文の転載、ログイン用URL／ワンタイムトークン、不要な個人識別情報は複製しない。
- 予定の`importance`等は不要な変更をしない。`sensitivity: private`を推奨。ただし共有権限や組織管理者のアクセスを完全に遮断する機能ではない。
- 確定は`showAs: busy`、仮予定は`showAs: tentative`＋件名`【仮】`を初期案とする。
- 成功通知は送らない。リマインダーは仮予定OFF、確定予定30分前などを**PoCで調整**。

### 9.2 予約の照合と状態管理

- 予約番号があれば優先的に利用。なければ送信元、施設名、利用者、旧日時、サービス内容、メールスレッドなどから一意性を検証。
- **自動作成した予定のみが自動更新・取消の対象**。独立した管理情報（予約キー ↔ Outlook Event ID）をAzure Tables等へ保持。
- Graphの`transactionId`（対応する操作・条件を実装時に検証）と独自予約キーで再試行時の重複登録を抑制する。
- ユーザーが手動編集／削除したイベントは上書き・復活させない方針を基本とし、PoCで競合処理を確定。
- 元予約と対応付けられない取消・変更メールは**既存予定に影響させず保留**。
- OutlookがICS招待から登録済みの予約について、二重作成しないよう既存予定を参照する。

### 9.3 キャンセル反映（未決定）

- 現行推奨は**即削除せず、件名に`【取消】`を付ける可逆的な状態変更**。
- 取消済みイベントの`showAs`を`free`にするか、日付履歴として維持するかはPoCで決定。
- 実際に予定を削除する方式を採る場合は、復元・監査・誤照合の観点で別途承認が必要。

## 10. 状態保存と冪等性

Azure Tables互換の永続ストアを初期候補とする。

| 論理テーブル | 保存対象 | 備考 |
|---|---|---|
| `SyncState` | Inbox `deltaLink`、初期同期完了日時、世代 | **PoCと本番は別物** |
| `ProcessedMessages` | Immutable Message ID、Internet Message ID、処理状態、受信日時 | 再処理・二重処理を防止 |
| `Bookings` | 予約キー、予約番号、施設、日時、状態、Event ID | 管理対象イベントの同一性 |
| `Operations` | イベント操作の冪等キー、結果、再試行、保留理由 | 障害復旧・診断 |
| `ModelUsage`（任意） | 推論回数、token数、モデル、処理時間 | **メール原文を保存しない** |

初回同期は**差分同期の基準点を確立**して既存メールの登録を抑制する。変更通知・重複実行・タイムアウト時でも「同一予約を1つの予定に収束」させることを目標とする。

## 11. モデル選定とFoundryの確認事項

モデルは**既存Foundryのデプロイを確認するまで未確定**。`model-router`（BalancedまたはCostモード）は比較候補であり必須要件ではない。

**評価軸：** 日本語予約文面の分類、日時精度、仮予約判定、変更・取消の照合、Structured Outputsへの対応、処理遅延、token使用量、単価。

- 固定モデルとmodel-routerの比較は、**既存Foundryで使えるもの**を優先。
- モデル出力だけで外部操作を決定せず、日時の根拠・予約同一性・更新対象IDをアプリ側で検証。
- `model-router`が新しい全モデルを自動的に利用できるとは限らず、対応モデルやバージョンはデプロイ時に確認する。
- 国内外の処理ロケーションはGlobalを許容済み。
- メール全件が推論対象とは限らない。決定論的な明らかなノイズ排除も検討するが、**予約の見逃しにつながる過剰なフィルタは避ける**。

## 12. 安全性・データ保護

- 医療情報のカレンダー本文保存は承諾済み。ただし機微情報であるため、アクセス権限と共有範囲、診断ログの収集設定を確認。
- **メール本文・AIへの入力全文・予約の医療詳細・Graph／Foundryの認証ヘッダーを通常ログに記録しない**。
- Foundryの外部推論呼び出し回数・トークン量・エラー率を計測し、異常時は停止できる上限を設ける。
- HTMLやURLなどメール由来データは信頼しない。メール本文に書かれた指示でシステム設定・外部送信先・権限を変更しない。
- Graph権限は対象メールボックスに限定し、カレンダーへの操作は**管理対象予定のみに限定**。
- 成功通知は送らないが、重大障害・認証期限切れ・連続失敗・保留案件の運用通知方式はPoC中に決める。
- ローカルSecretファイル／本番設定値に対するアクセス権限を絞り、シークレットをログに出さない。

## 13. PoCと開発の手順

| フェーズ | 内容 | 予定表書き込み | 通過条件 |
|---|---|---|---|
| 0. 接続準備 | Graphアプリ登録、Client Secret、Exchange RBAC、既存Foundry接続情報、Azurite | **なし** | ローカルからGraph対象メール取得とFoundry架空入力の疎通成功。対象外Graph拒否 |
| 1. 抽出PoC | 合成メール＋許可された実メール20〜50通の予約判定・日時抽出 | **なし** | 重大な日時誤抽出・無関係メールの誤登録候補を洗い出す |
| 2. 仕様確定 | Structured Outputs、バリデーション、状態遷移、仮予定基準 | **なし** | テストケース追加・再実行で挙動安定 |
| 3. 書き込み試験 | 明示承認されたテスト予定で作成・変更・取消・冪等性 | **テストイベントのみ** | 手動既存予定の変更なし、重複作成なし |
| 4. 新着監視 | Delta基準点作成、ページング、カーソル永続化、5分Timer | **`dry-run`** | 再起動・再試行で取りこぼし・二重登録がない |
| 5. 本番導入 | 専用Functionsをデプロイ、観測後書き込みモードを順次有効化 | **個別承認後** | 運用監視と停止手順が機能 |
| 6. 最適化 | モデル比較、費用計測、通知・添付対応 | 設定どおり | 精度・費用・保守の改善 |

**受け入れ基準（暫定）：** テスト対象内で重大な日時誤登録候補0件、無関係メールの作成候補0件、管理外イベントへの変更0件、再実行時の二重作成0件。50件で0件でも将来の無事故を保証しないため、本番投入後も測定を続ける。

## 14. 人間が準備するもの／作業を依頼するタイミング

### 14.1 いま確認・準備すべきもの

1. **Exchange Online／Entra IDの管理者サインインが可能であること。** Exchange管理者権限はユーザー申告済み。Application RBAC用の`Organization Management`等の委任権限は実行時に検証。
2. **Graph用の専用Entra IDアプリ登録を作成**し、`Client ID`・`Tenant ID`を取得する。PoCと本番で同じアプリ登録を使うことを推奨。
3. **Graph Client Secretを作成**し、秘密値をローカル開発環境とFunctions本番アプリ設定へ安全に投入する。開発用・本番用に別Secretを発行してもよい（推奨）。**Secret値は文書・チャットへ投稿しない。**
4. Exchange Online Application RBACで対象メールボックスを限定する。管理者MFAが必要な場面はユーザーが対応する。
5. **別テナントFoundryのEndpoint、Deployment名、利用可能なAPI形式、APIキーを確認**する。**APIキー値はチャットへ貼らない。**
6. FoundryのAPIキー認証が有効か、Functionsから公開HTTPSでアクセス可能か、推論課金サブスクリプションが想定どおりか確認する。

### 14.2 開発プロセス内で構築可能なもの（承認後）

- 専用Function App（Japan East、Flex Consumption、Always Ready無効）。
- Storage／必要最小限のログ設定／Bicepによる再現可能なIaC。
- Graph／Foundry統合クライアント、ローカルCLI、Azurite接続、予約判定・状態管理。
- Azureの予算アラートとアプリ側の安全な推論回数ガード。
- GitHub CI、テスト、デプロイ設定（Secretがログやリポジトリに出ない運用）。

### 14.3 明示承認が必要な操作

- 新しいAzureリソースの作成、既存Azure設定の変更、実際の課金が伴うモデル呼び出し。
- Exchange Onlineでのアプリ登録・Secret生成・RBAC変更（設定実施の了承）。
- **メインカレンダーへのイベント作成・変更・取消**。PoCでは原則`dry-run`。
- 既存Foundryで新たにモデルをデプロイするなどの変更。

## 15. Azure実環境で確認できたこと（2026-10-17時点）

### 15.1 既存Function Apps

| Function App | プラン | 場所 | 確認時の稼働状態 |
|---|---|---|---|
| `func-discord-reader-mcp` | Flex Consumption（FC1） | Japan East | Running |
| `func-watch-wishlist-sale` | Consumption（Y1） | Japan East | Running |
| `func-github-webhook-bridge` | Consumption（Y1） | Japan East | Running |

- すべて`Main Azure subscription`に存在。PineのAzure CLIで同サブスクリプションのOwner権限が確認された。
- 別テナントの既存Foundryはメイン側CLIから参照できない。**Main側に新規Foundryは作成しない**。
- 2026年9月の既存3アプリ関連リソースグループのAzure Cost Management `ActualCost`は**合計約613円**。Functions計算課金は**0円**、残額は主にStorageとAzure Monitor。
- **これは既存システムの請求額であり、今回のメール自動登録システムの月額確定額ではない**。

### 15.2 メール実測（現在残存するInbox＋Deleted Items）

| 月 | Inbox | Deleted Items | 合計 |
|---|---:|---:|---:|
| 2026年7月 | 0 | 409 | 409 |
| 2026年8月 | 0 | 356 | 356 |
| 2026年9月 | 1 | 396 | 397 |
| 2026年10月17日まで | 7 | 356 | 363 |
| **7〜9月の平均** | | | **約387通／月** |

- 初期容量設計は**400通／月**、余裕を考慮すると500通／月。
- Inboxとゴミ箱の残存数の合計であり、当時Inboxを通過した全新着数とは一致しない。完全削除済みメールも含まれない。
- **Inboxのみを監視し、ゴミ箱のメールは処理対象に追加しない。**

### 15.3 月額費用の見通し

- **Azure Functions：** 既存プランの無料枠に収まれば**計算料金0円**が見込める。ただし、Storage・監視・ネットワーク等は別料金。
- **Foundry：** **既存Foundry側サブスクリプションに課金**。モデル、Router設定、推論token数が確定しない限り金額は確約できない。
- PoCで**処理メール数／モデル呼び出し数／入力・出力token数／エラー・再試行数**を収集し、実請求と突き合わせる。
- 月額数ドル規模が当初の予算感だが、**上限保証ではない**。予算アラートは費用を自動停止しない。

## 16. 未確定事項とブロッカー

| 事項 | 段階 | 判断・検証内容 |
|---|---|---|
| **Graphアプリ登録とSecret** | 接続PoC前・必須 | 新規専用アプリ、Secret期限、PoC／本番の資格情報配置 |
| **Exchange Application RBAC** | Graph接続前・必須 | 管理者権限、Mailbox Scope、実Graphでの対象内・対象外アクセス |
| **Foundry接続情報** | 推論PoC前・必須 | Endpoint、Deployment、API形式、キー、ネットワーク到達性 |
| **PoCで実メールをFoundryへ送信する範囲** | 実メール評価前・要確認 | ユーザーが選んだサンプルのみか、新着をdry-run解析するか |
| **PoC書き込みテストの相手先** | Graph書き込みテスト前・要確認 | テスト予定のみ／テスト用メールボックス等。メイン予定への試験書き込みには承認 |
| **本番カレンダー書き込み** | 本番化前・必須承認 | `create-only`／`full`の切替、停止・復旧手順 |
| **取消方法** | 取消自動反映前 | `【取消】`＋freeで保存／削除。現行推奨は保存 |
| **手動編集・削除への対処** | 更新有効化前 | ユーザー優先、再生成抑止、楽観ロック等 |
| **通知・リマインダー** | PoCで調整 | 仮予定・確定予定の通知、例外連絡先 |
| **モデル最終選択** | PoCで決定 | 既存Foundryデプロイ、精度、Structured Outputs、実料金 |
| **添付対応** | 後続でも可 | ICS重複判定、PDF・画像対応の優先順位 |
| **Storageの具体的な認証方式** | デプロイ時 | ホスト・デプロイ用接続とアプリ状態用接続を確定 |
| **Secretローテーション運用** | 本番化前 | 所有者・期限・差し替えと復旧手順 |

## 17. 実装開始までの作業順序（推奨）

1. **このv0.2文書をレビュー**して、Graph Client Secret方式とテスト書き込みポリシーを承認する。
2. **Entraの専用Graphアプリを登録・Client Secretを発行**し、Exchange Application RBACを安全に設定する。
3. **既存FoundryのEndpoint・DeploymentとAPIキーを安全に設定**する（別テナント側に新規リソースは作らない）。
4. **PineでローカルCLI + AzuriteのPoCを実装**し、Graphの読み取りとFoundryの架空メール推論を試験する。
5. 合成メール・許可された実メールで分類と日時検証を繰り返し、API使用量を計測する。
6. 書き込みテストの承認後、管理対象テストイベントの作成・変更・取消を検証する。
7. 専用Azure Functionsを構築・デプロイし、**本番は`dry-run`から開始**する。
8. ユーザーによる明示承認後、メインカレンダーへの`create-only`、続いて安全な`full`を段階的に有効化する。

## 18. 参照資料（Microsoft公式）

- [Microsoft identity platform：OAuth 2.0 Client Credentials](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-client-creds-grant-flow)
- [Graphアプリ専用認証の設定ガイド](https://learn.microsoft.com/graph/tutorials/dotnet-app-only)
- [Exchange Online：Role Based Access Control for Applications](https://learn.microsoft.com/exchange/permissions-exo/application-rbac)
- [Microsoft Graph：Message Delta](https://learn.microsoft.com/graph/api/message-delta)
- [Graph：フォルダー内メール差分追跡](https://learn.microsoft.com/graph/delta-query-messages)
- [Azure Functions：ローカル開発](https://learn.microsoft.com/azure/azure-functions/functions-develop-local)
- [Azure Functions：Timer Trigger](https://learn.microsoft.com/azure/azure-functions/functions-bindings-timer)
- [Azure Functions：Flex Consumption](https://learn.microsoft.com/azure/azure-functions/flex-consumption-how-to)
- [Azure Storage：Azurite](https://learn.microsoft.com/azure/storage/common/storage-use-azurite)
- [Foundry：REST API認証](https://learn.microsoft.com/azure/foundry/openai/reference)
- [Foundry：モデルルーター](https://learn.microsoft.com/azure/foundry/openai/how-to/model-router)
- [Foundry：Structured Outputs](https://learn.microsoft.com/azure/foundry/openai/how-to/structured-outputs)
- [Azure Functions：従量課金コスト](https://learn.microsoft.com/azure/azure-functions/functions-consumption-costs)

---

**取り扱い上の注意：** 本書は技術仕様とPoC計画の合意用資料。APIキー／Client Secret／アクセストークンなどの値は保存しない。設定変更・実課金の発生する操作・メインカレンダー書き込み・既存Foundryへの変更は、ユーザーの明示的承認を必要とする。

