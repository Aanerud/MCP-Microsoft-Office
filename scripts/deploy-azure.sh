#!/bin/bash

# MCP Microsoft Office - Full Azure Deployment Script
# This script creates all Azure infrastructure from scratch

set -e

# Configuration
RESOURCE_GROUP="Frutti"
LOCATION="norwayeast"
APP_SERVICE_PLAN="mcp-nstop-plan"
WEB_APP_NAME="mcp-nstop"
RUNTIME="NODE:20-lts"
SKU="B1"
CUSTOM_DOMAIN="mcp.nstop.no"

# Microsoft Azure AD Configuration (from .env.production)
MICROSOFT_CLIENT_ID="6c9b2994-b11f-4a62-83a2-219210cc927c"
MICROSOFT_TENANT_ID="facf3dd7-5092-42e2-bea6-1a201d58b8f6"

echo "============================================"
echo "MCP Microsoft Office - Azure Deployment"
echo "============================================"
echo ""

# Check if logged in to Azure
echo "Checking Azure CLI login status..."
if ! az account show &>/dev/null; then
    echo "ERROR: Not logged in to Azure CLI. Run 'az login' first."
    exit 1
fi

ACCOUNT_NAME=$(az account show --query name -o tsv)
echo "Logged in to: $ACCOUNT_NAME"
echo ""

# Step 1: Create Resource Group
echo "Step 1: Creating Resource Group..."
if az group show --name $RESOURCE_GROUP &>/dev/null; then
    echo "  Resource group '$RESOURCE_GROUP' already exists."
else
    az group create --name $RESOURCE_GROUP --location $LOCATION --output none
    echo "  Created resource group '$RESOURCE_GROUP' in $LOCATION"
fi
echo ""

# Step 2: Create App Service Plan
echo "Step 2: Creating App Service Plan..."
if az appservice plan show --name $APP_SERVICE_PLAN --resource-group $RESOURCE_GROUP &>/dev/null; then
    echo "  App Service Plan '$APP_SERVICE_PLAN' already exists."
else
    az appservice plan create \
        --name $APP_SERVICE_PLAN \
        --resource-group $RESOURCE_GROUP \
        --sku $SKU \
        \
        --output none
    echo "  Created App Service Plan '$APP_SERVICE_PLAN' (SKU: $SKU, Linux)"
fi
echo ""

# Step 3: Create Web App
echo "Step 3: Creating Web App..."
if az webapp show --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP &>/dev/null; then
    echo "  Web App '$WEB_APP_NAME' already exists."
else
    az webapp create \
        --name $WEB_APP_NAME \
        --resource-group $RESOURCE_GROUP \
        --plan $APP_SERVICE_PLAN \
        --runtime "$RUNTIME" \
        --output none
    echo "  Created Web App '$WEB_APP_NAME' (Runtime: $RUNTIME)"
fi
echo ""

# Step 4: Generate Secrets
echo "Step 4: Generating Secrets..."
STATIC_JWT_SECRET=$(openssl rand -base64 32)
JWT_SECRET=$(openssl rand -base64 32)
# Device registry key must be exactly 32 bytes for AES-256
DEVICE_REGISTRY_KEY=$(openssl rand -hex 16)

echo "  Generated new JWT secrets and encryption key"
echo ""

# Step 5: Configure Environment Variables
echo "Step 5: Configuring Environment Variables..."
az webapp config appsettings set \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --settings \
        MICROSOFT_CLIENT_ID="$MICROSOFT_CLIENT_ID" \
        MICROSOFT_TENANT_ID="$MICROSOFT_TENANT_ID" \
        MICROSOFT_REDIRECT_URI="https://$CUSTOM_DOMAIN/api/auth/callback" \
        STATIC_JWT_SECRET="$STATIC_JWT_SECRET" \
        JWT_SECRET="$JWT_SECRET" \
        MCP_BEARER_TOKEN_EXPIRY="24h" \
        DEVICE_REGISTRY_ENCRYPTION_KEY="$DEVICE_REGISTRY_KEY" \
        NODE_ENV="production" \
        HOST="0.0.0.0" \
        PORT="8080" \
        SERVER_URL="https://$CUSTOM_DOMAIN/" \
        DOMAIN="$CUSTOM_DOMAIN" \
        LLM_PROVIDER="openai" \
        CORS_ALLOWED_ORIGINS="https://$CUSTOM_DOMAIN,https://$WEB_APP_NAME.azurewebsites.net" \
        WEBSITE_NODE_DEFAULT_VERSION="~20" \
        SCM_DO_BUILD_DURING_DEPLOYMENT="true" \
    --output none

echo "  Configured all environment variables"
echo ""

# Step 6: Configure Web App settings
echo "Step 6: Configuring Web App settings..."
az webapp config set \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --startup-file "npm start" \
    --output none

az webapp update \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --https-only true \
    --output none

echo "  Configured startup command and HTTPS-only"
echo ""

# Step 7: Get Publish Profile for GitHub Actions
echo "Step 7: Getting Publish Profile..."
PUBLISH_PROFILE=$(az webapp deployment list-publishing-profiles \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --xml)

echo "  Publish profile retrieved"
echo ""

# Save publish profile to file (temporarily)
PROFILE_FILE="/tmp/publish-profile-$WEB_APP_NAME.xml"
echo "$PUBLISH_PROFILE" > "$PROFILE_FILE"
echo "  Saved to: $PROFILE_FILE"
echo ""

# Step 8: Display DNS Configuration Required
echo "============================================"
echo "DNS CONFIGURATION REQUIRED"
echo "============================================"
echo ""
echo "Before adding custom domain, ensure these DNS records exist:"
echo ""
echo "  Type: CNAME"
echo "  Name: mcp"
echo "  Value: $WEB_APP_NAME.azurewebsites.net"
echo ""

# Get the custom domain verification ID
VERIFICATION_ID=$(az webapp show --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP --query "customDomainVerificationId" -o tsv 2>/dev/null || echo "")
if [ -n "$VERIFICATION_ID" ]; then
    echo "  Type: TXT"
    echo "  Name: asuid.mcp"
    echo "  Value: $VERIFICATION_ID"
    echo ""
fi

echo "============================================"
echo "DEPLOYMENT COMPLETE"
echo "============================================"
echo ""
echo "Azure Default URL: https://$WEB_APP_NAME.azurewebsites.net"
echo ""
echo "NEXT STEPS:"
echo "1. Verify DNS records are configured correctly"
echo "2. Run: ./scripts/add-custom-domain.sh"
echo "3. Update GitHub secret 'AZURE_WEBAPP_PUBLISH_PROFILE' with contents of:"
echo "   $PROFILE_FILE"
echo "4. Push to main branch to trigger GitHub Actions deployment"
echo ""
echo "To update GitHub secret manually:"
echo "  gh secret set AZURE_WEBAPP_PUBLISH_PROFILE < $PROFILE_FILE"
echo ""
