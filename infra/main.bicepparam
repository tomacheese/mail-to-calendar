using 'main.bicep'

param graphTenantId = readEnvironmentVariable('GRAPH_TENANT_ID')
param graphClientId = readEnvironmentVariable('GRAPH_CLIENT_ID')
param graphClientSecret = readEnvironmentVariable('GRAPH_CLIENT_SECRET')
param graphMailbox = readEnvironmentVariable('GRAPH_MAILBOX')
param foundryEndpoint = readEnvironmentVariable('FOUNDRY_ENDPOINT')
param foundryDeployment = readEnvironmentVariable('FOUNDRY_DEPLOYMENT')
param foundryApiKey = readEnvironmentVariable('FOUNDRY_API_KEY')
param calendarWriteMode = readEnvironmentVariable('CALENDAR_WRITE_MODE', 'dry-run')
