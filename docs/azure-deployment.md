# Deploying to Azure App Service

This guide covers deploying the MCP Microsoft Office server to Azure App Service using GitHub Actions.

## Prerequisites

- Azure account with an active subscription
- GitHub repository with the project code
- Azure CLI installed locally (for initial setup)

## Architecture Overview

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  GitHub Repo    │────►│  GitHub Actions │────►│  Azure App      │
│  (main branch)  │     │  (CI/CD)        │     │  Service        │
└─────────────────┘     └─────────────────┘     └─────────────────┘
                                                        │
                                                        ▼
                                                ┌─────────────────┐
                                                │  Microsoft 365  │
                                                │  Graph API      │
                                                └─────────────────┘
```

## Step 1: Create Azure Resources

### Using Azure CLI

```bash
# Login to Azure
az login

# Set variables
RESOURCE_GROUP="your-resource-group"
WEB_APP_NAME="your-app-name"
LOCATION="norwayeast"  # or your preferred region

# Create resource group
az group create --name $RESOURCE_GROUP --location $LOCATION

# Create App Service Plan (B1 tier recommended for production)
az appservice plan create \
    --name "${WEB_APP_NAME}-plan" \
    --resource-group $RESOURCE_GROUP \
    --sku B1 \
    --is-linux

# Create Web App with Node.js 20
az webapp create \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --plan "${WEB_APP_NAME}-plan" \
    --runtime "NODE|20-lts"
```

## Step 2: Configure App Settings

Set required environment variables in Azure:

```bash
az webapp config appsettings set \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --settings \
    NODE_ENV="production" \
    PORT="8080" \
    MICROSOFT_CLIENT_ID="your-client-id" \
    MICROSOFT_TENANT_ID="your-tenant-id" \
    MICROSOFT_REDIRECT_URI="https://${WEB_APP_NAME}.azurewebsites.net/api/auth/callback" \
    JWT_SECRET="your-secure-jwt-secret-at-least-64-chars" \
    DEVICE_REGISTRY_ENCRYPTION_KEY="your-32-byte-encryption-key" \
    CORS_ALLOWED_ORIGINS="https://${WEB_APP_NAME}.azurewebsites.net"
```

### Important Settings

| Setting | Description |
|---------|-------------|
| `PORT` | Must be `8080` - Azure expects this port |
| `NODE_ENV` | Set to `production` |
| `MICROSOFT_REDIRECT_URI` | Must match Azure App Registration |
| `SCM_DO_BUILD_DURING_DEPLOYMENT` | Set to `false` if using pre-built artifacts |

## Step 3: Configure Startup Command

Azure needs a custom startup command to handle the deployment correctly:

```bash
az webapp config set \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --startup-file "bash startup.sh"
```

The project includes a `startup.sh` file that:
1. Extracts node_modules from the compressed tar.gz format
2. Sets up the correct NODE_PATH
3. Starts the server with proper environment variables

## Step 4: Set Up GitHub Actions

### Get Publish Profile

```bash
az webapp deployment list-publishing-profiles \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --xml
```

### Add GitHub Secret

1. Go to your GitHub repository
2. Navigate to **Settings** → **Secrets and variables** → **Actions**
3. Add a new secret named `AZURE_WEBAPP_PUBLISH_PROFILE`
4. Paste the entire XML output from the previous command

### GitHub Actions Workflow

The project includes `.github/workflows/azure-deploy.yml` which:
- Triggers on push to `main` branch
- Installs production dependencies only (`npm ci --omit=dev`)
- Uploads the entire project as an artifact
- Deploys to Azure using the publish profile

## Step 5: Update Azure App Registration

Add the Azure callback URL to your Azure App Registration:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** → **App registrations**
3. Select your app
4. Go to **Authentication**
5. Add redirect URI: `https://your-app-name.azurewebsites.net/api/auth/callback`

## Troubleshooting

### Startup Probe Timeout (230 seconds)

**Symptom:** App shows "Application Error" and logs show startup probe failed.

**Cause:** Node.js modules take too long to load before the HTTP server starts listening.

**Solution:** The server is designed to start an HTTP server immediately at module load time, before loading heavy dependencies. This ensures Azure's startup probe passes within seconds.

Key code pattern in `src/main/dev-server.cjs`:
```javascript
// Start server IMMEDIATELY at module load time
const earlyApp = express();
earlyApp.get('/api/health', (req, res) => res.json({ status: 'starting' }));
earlyApp.listen(PORT, HOST, () => {
  console.log(`Early health endpoint running on ${HOST}:${PORT}`);
});

// THEN load heavy modules (these can take time)
const heavyModule = require('./heavy-module.cjs');
```

### node_modules Not Found

**Symptom:** App crashes with "Cannot find module" errors.

**Cause:** Azure's Oryx build system compresses node_modules to `node_modules.tar.gz`.

**Solution:** The `startup.sh` script handles extraction:
```bash
if [ -f node_modules.tar.gz ]; then
    tar -xzf node_modules.tar.gz -C /node_modules
    export NODE_PATH="/node_modules:$NODE_PATH"
    ln -sfn /node_modules ./node_modules
fi
```

### Logs Not Showing Recent Code

**Symptom:** Console.log statements don't appear after deployment.

**Cause:** Azure doesn't automatically restart the container after deployment.

**Solution:** Restart the web app after deployment:
```bash
az webapp restart --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP
```

Or enable automatic restart in the deployment workflow.

### Viewing Logs

**Real-time log streaming:**
```bash
az webapp log tail --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP
```

**Download logs:**
```bash
az webapp log download --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP --log-file logs.zip
```

**Enable container logging:**
```bash
az webapp log config \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --docker-container-logging filesystem
```

### Health Check Endpoints

The server provides these endpoints for monitoring:

| Endpoint | Purpose |
|----------|---------|
| `/api/health` | Basic health check - returns `{"status":"ok"}` |
| `/api/status` | Detailed status including auth state |

Configure Azure health check:
```bash
az webapp config set \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --generic-configurations '{"healthCheckPath": "/api/health"}'
```

## Production Checklist

Before going live:

- [ ] Set `NODE_ENV=production`
- [ ] Configure `JWT_SECRET` (64+ characters)
- [ ] Configure `DEVICE_REGISTRY_ENCRYPTION_KEY` (32 bytes)
- [ ] Set `CORS_ALLOWED_ORIGINS` to your domain
- [ ] Update Azure App Registration with production redirect URI
- [ ] Enable HTTPS only: `az webapp update --https-only true`
- [ ] Configure custom domain (optional)
- [ ] Set up monitoring and alerts in Azure Portal

## Cost Optimization

- **B1 tier** (~$13/month): Recommended for small teams
- **F1 tier** (free): Limited, good for testing only
- **P1V2 tier** (~$80/month): For high traffic production use

## Security Considerations

1. **Never commit secrets** - Use Azure App Settings or Key Vault
2. **Use managed identity** for Azure resources when possible
3. **Enable Azure AD authentication** for admin access
4. **Review firewall rules** - Restrict access if needed
5. **Enable backup** for the SQLite database file in `/home/data/`

## Useful Commands

```bash
# Check app status
az webapp show --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP --query state

# View app settings
az webapp config appsettings list --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP

# Scale up/down
az appservice plan update --name "${WEB_APP_NAME}-plan" --resource-group $RESOURCE_GROUP --sku B2

# Stop/start app
az webapp stop --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP
az webapp start --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP
```
