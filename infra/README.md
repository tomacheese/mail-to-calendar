# infra

Bicep template deploying `mail-to-calendar` as an Azure Functions v4 app
(Node.js 24 LTS, Flex Consumption / FC1, no Always Ready instances), into an
**existing** resource group:

- Log Analytics Workspace (daily ingestion cap 0.1 GB)
- Application Insights (workspace-based)
- Storage Account (also holds the Flex Consumption deployment package container and the app state tables)
- App Service Plan (Linux, Flex Consumption / FC1)
- Function App (Linux, Node.js 24 runtime, max 1 instance so timer runs never race on the delta cursor)

The Foundry resource is **not** created here: the app uses an existing one in another tenant via API key.

## Deploy

Creating resources and billing start here, so run it only after explicit approval.

```bash
az group create --name rg-mail-to-calendar --location japaneast

cd infra
set -a && . ../.env && set +a   # or export the variables below yourself
az deployment group create \
  --resource-group rg-mail-to-calendar \
  --template-file main.bicep \
  --parameters main.bicepparam
```

`main.bicepparam` reads the values from environment variables:

| Variable                                              | Secret |
| ----------------------------------------------------- | ------ |
| `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_MAILBOX` | -      |
| `GRAPH_CLIENT_SECRET`                                 | yes    |
| `FOUNDRY_ENDPOINT`, `FOUNDRY_DEPLOYMENT`              | -      |
| `FOUNDRY_API_KEY`                                     | yes    |
| `CALENDAR_WRITE_MODE` (optional, default `dry-run`)   | -      |

The first deployment must stay on `dry-run`. Switching to `create-only` / `full` writes to the main calendar and needs explicit approval.

Application code is deployed separately via
[`.github/workflows/azure-functions-deploy.yml`](../.github/workflows/azure-functions-deploy.yml)
(push to `main`/`master`, or manual `workflow_dispatch`) — this template only
provisions the infrastructure.

**If deploying manually** (`func azure functionapp publish` / `az functionapp
deploy`), install prod dependencies with
`pnpm install --prod --frozen-lockfile --config.node-linker=hoisted` first.
pnpm's default symlinked `node_modules` does not survive Kudu's zip
extraction — transitive deps resolve fine locally but come back "Cannot find
module" once deployed.

## Required GitHub Secrets

| Secret Name              | Purpose                                                            |
| ------------------------ | ------------------------------------------------------------------ |
| `AZURE_CLIENT_ID`        | Client ID of the Azure AD application used for OIDC authentication |
| `AZURE_TENANT_ID`        | Azure AD tenant ID                                                 |
| `AZURE_SUBSCRIPTION_ID`  | Azure subscription ID of the deployment target                     |
| `AZURE_FUNCTIONAPP_NAME` | Name of the Function App to deploy to (`func-mail-to-calendar`)    |

These are expected to be configured in the `production` environment of the
GitHub repository (see `environment: production` in the workflow).
