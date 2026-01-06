#!/bin/bash

# MCP Microsoft Office - Add Custom Domain and SSL
# Run this after DNS records are configured

set -e

# Configuration
RESOURCE_GROUP="mcp-microsoft-office-rg"
WEB_APP_NAME="mcp-microsoft-office"
CUSTOM_DOMAIN="mcp.nstop.no"

echo "============================================"
echo "Adding Custom Domain and SSL Certificate"
echo "============================================"
echo ""

# Check if webapp exists
if ! az webapp show --name $WEB_APP_NAME --resource-group $RESOURCE_GROUP &>/dev/null; then
    echo "ERROR: Web App '$WEB_APP_NAME' not found. Run deploy-azure.sh first."
    exit 1
fi

# Step 1: Verify DNS
echo "Step 1: Checking DNS configuration..."
echo "  Verifying CNAME for $CUSTOM_DOMAIN..."

# Check CNAME
CNAME_CHECK=$(dig +short CNAME $CUSTOM_DOMAIN 2>/dev/null || echo "")
if [[ "$CNAME_CHECK" == *"$WEB_APP_NAME.azurewebsites.net"* ]]; then
    echo "  CNAME record is correctly configured"
else
    echo "  WARNING: CNAME may not be configured correctly"
    echo "  Expected: $WEB_APP_NAME.azurewebsites.net"
    echo "  Found: $CNAME_CHECK"
    echo ""
    read -p "Continue anyway? (y/N) " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
fi
echo ""

# Step 2: Add Custom Domain
echo "Step 2: Adding custom domain..."
if az webapp config hostname list --webapp-name $WEB_APP_NAME --resource-group $RESOURCE_GROUP --query "[?name=='$CUSTOM_DOMAIN']" -o tsv | grep -q "$CUSTOM_DOMAIN"; then
    echo "  Custom domain '$CUSTOM_DOMAIN' already configured"
else
    az webapp config hostname add \
        --webapp-name $WEB_APP_NAME \
        --resource-group $RESOURCE_GROUP \
        --hostname $CUSTOM_DOMAIN \
        --output none
    echo "  Added custom domain '$CUSTOM_DOMAIN'"
fi
echo ""

# Step 3: Create Managed SSL Certificate
echo "Step 3: Creating managed SSL certificate..."
az webapp config ssl create \
    --name $WEB_APP_NAME \
    --resource-group $RESOURCE_GROUP \
    --hostname $CUSTOM_DOMAIN \
    --output none 2>/dev/null || echo "  Certificate may already exist or is being provisioned"
echo "  SSL certificate created/requested"
echo ""

# Step 4: Bind SSL Certificate
echo "Step 4: Binding SSL certificate..."
THUMBPRINT=$(az webapp config ssl list --resource-group $RESOURCE_GROUP --query "[?subjectName=='$CUSTOM_DOMAIN' || contains(subjectName, '$CUSTOM_DOMAIN')].thumbprint" -o tsv | head -1)

if [ -n "$THUMBPRINT" ]; then
    az webapp config ssl bind \
        --name $WEB_APP_NAME \
        --resource-group $RESOURCE_GROUP \
        --certificate-thumbprint "$THUMBPRINT" \
        --ssl-type SNI \
        --output none 2>/dev/null || echo "  SSL may already be bound"
    echo "  SSL certificate bound to domain"
else
    echo "  Certificate is being provisioned. This may take a few minutes."
    echo "  Re-run this script in 5-10 minutes to complete SSL binding."
fi
echo ""

echo "============================================"
echo "CUSTOM DOMAIN SETUP COMPLETE"
echo "============================================"
echo ""
echo "Your site should be available at:"
echo "  https://$CUSTOM_DOMAIN"
echo ""
echo "Note: SSL certificate provisioning may take 5-15 minutes."
echo ""
