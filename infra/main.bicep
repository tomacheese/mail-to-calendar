// Azure Function App: mail-to-calendar
// Deploy target: existing resource group (this template does not create the resource group itself)
// Runtime: Azure Functions v4, Node.js 24 LTS, Flex Consumption (FC1), no Always Ready instances
@description('Azure region for all resources')
param location string = resourceGroup().location

@description('Tenant ID of the Entra ID tenant that owns the target mailbox')
param graphTenantId string

@description('Application (client) ID of the Graph app registration')
param graphClientId string

@description('Client secret value of the Graph app registration')
@secure()
param graphClientSecret string

@description('UPN or ID of the mailbox to watch')
param graphMailbox string

@description('Endpoint of the existing Foundry resource (other tenant)')
param foundryEndpoint string

@description('Name of the existing Foundry model deployment')
param foundryDeployment string

@description('API key of the existing Foundry resource')
@secure()
param foundryApiKey string

@description('Calendar write mode. Starts at dry-run; switching to create-only/full needs explicit approval')
@allowed([
  'dry-run'
  'create-only'
  'full'
])
param calendarWriteMode string = 'dry-run'

@description('Upper bound of Foundry calls per run')
param foundryMaxCallsPerRun int = 20

var logAnalyticsName = 'log-mail-to-calendar'
var appInsightsName = 'appi-mail-to-calendar'
var storageAccountName = 'stmailtocalendar'
var appServicePlanName = 'asp-mail-to-calendar-flex'
var functionAppName = 'func-mail-to-calendar'
var deploymentContainerName = 'app-package-${functionAppName}'

resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: logAnalyticsName
  location: location
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    retentionInDays: 30
    workspaceCapping: {
      // Cost guard: a 5-minute timer must not turn into an ingestion bill.
      dailyQuotaGb: json('0.1')
    }
  }
}

resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: appInsightsName
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logAnalytics.id
    IngestionMode: 'LogAnalytics'
  }
}

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: storageAccountName
  location: location
  kind: 'StorageV2'
  sku: {
    name: 'Standard_LRS'
  }
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
  }
}

resource blobService 'Microsoft.Storage/storageAccounts/blobServices@2023-05-01' existing = {
  parent: storageAccount
  name: 'default'
}

// Flex Consumption deploys code from a blob container instead of the classic
// WEBSITE_RUN_FROM_PACKAGE / content-share model.
resource deploymentContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = {
  parent: blobService
  name: deploymentContainerName
}

resource appServicePlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: appServicePlanName
  location: location
  kind: 'linux'
  sku: {
    name: 'FC1'
    tier: 'FlexConsumption'
  }
  properties: {
    reserved: true
  }
}

var storageConnectionString = 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=${environment().suffixes.storage}'

resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: functionAppName
  location: location
  kind: 'functionapp,linux'
  tags: {
    'hidden-link: /app-insights-resource-id': appInsights.id
  }
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    clientAffinityEnabled: false
    siteConfig: {
      ftpsState: 'FtpsOnly'
      minTlsVersion: '1.2'
      http20Enabled: false
      appSettings: [
        {
          name: 'AzureWebJobsStorage'
          value: storageConnectionString
        }
        {
          name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
          value: appInsights.properties.ConnectionString
        }
        {
          name: 'APP_ENV'
          value: 'production'
        }
        {
          name: 'GRAPH_TENANT_ID'
          value: graphTenantId
        }
        {
          name: 'GRAPH_CLIENT_ID'
          value: graphClientId
        }
        {
          name: 'GRAPH_CLIENT_SECRET'
          value: graphClientSecret
        }
        {
          name: 'GRAPH_MAILBOX'
          value: graphMailbox
        }
        {
          name: 'FOUNDRY_ENDPOINT'
          value: foundryEndpoint
        }
        {
          name: 'FOUNDRY_DEPLOYMENT'
          value: foundryDeployment
        }
        {
          name: 'FOUNDRY_API_KEY'
          value: foundryApiKey
        }
        {
          name: 'FOUNDRY_MAX_CALLS_PER_RUN'
          value: string(foundryMaxCallsPerRun)
        }
        {
          // State lives in the app's own storage account; never shared with local Azurite.
          name: 'STATE_STORAGE_CONNECTION'
          value: storageConnectionString
        }
        {
          name: 'CALENDAR_WRITE_MODE'
          value: calendarWriteMode
        }
        {
          name: 'MAIL_POLLING_ENABLED'
          value: 'true'
        }
        {
          name: 'PROCESS_EXISTING_MESSAGES'
          value: 'false'
        }
        {
          name: 'DEFAULT_EVENT_DURATION_MINUTES'
          value: '60'
        }
      ]
    }
    functionAppConfig: {
      deployment: {
        storage: {
          type: 'blobContainer'
          value: '${storageAccount.properties.primaryEndpoints.blob}${deploymentContainerName}'
          authentication: {
            type: 'StorageAccountConnectionString'
            storageAccountConnectionStringName: 'AzureWebJobsStorage'
          }
        }
      }
      scaleAndConcurrency: {
        // One poller: concurrent timer runs would race on the delta cursor.
        maximumInstanceCount: 1
        instanceMemoryMB: 512
      }
      runtime: {
        name: 'node'
        version: '24'
      }
    }
  }
}
