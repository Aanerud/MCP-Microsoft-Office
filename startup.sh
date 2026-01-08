#!/bin/bash
# Azure App Service startup script
# This script ensures the server starts immediately to pass Azure's startup probe

cd /home/site/wwwroot

# Set up node_modules if using tar.gz format
if [ -f node_modules.tar.gz ]; then
    echo "Found tar.gz based node_modules."
    rm -rf /node_modules
    mkdir -p /node_modules
    tar -xzf node_modules.tar.gz -C /node_modules
    export NODE_PATH="/node_modules:$NODE_PATH"
    export PATH="/node_modules/.bin:$PATH"
    if [ -d node_modules ]; then
        rm -rf node_modules
    fi
    ln -sfn /node_modules ./node_modules
    echo "Done extracting modules."
fi

# Ensure PORT is set
export PORT=${PORT:-8080}

# Start the application
echo "Starting MCP server on port $PORT..."
exec node src/main/dev-server.cjs
