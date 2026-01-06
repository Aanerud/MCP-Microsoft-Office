/**
 * @fileoverview UI Manager for MCP renderer process
 * Handles UI state management, rendering, and DOM operations
 */

// Import utilities
import { DOMSafe } from '../dom-safe.js';
import { UINotification } from '../ui-notification.js';
import IPCService from '../services/IPCService.js';

export class UIManager {
    constructor() {
        this.renderCallbacks = new Map();
        this.initialized = false;
        this.currentUser = null; // Add this line to store user information
    }

    /**
     * Initialize UI Manager
     */
    async init() {
        if (this.initialized) return;
        
        try {
            if (window.MonitoringService) {
                window.MonitoringService.info('Initializing UI Manager', { operation: 'ui-init' }, 'renderer');
            }
            this.setupEventListeners();
            this.initialized = true;
            if (window.MonitoringService) {
                window.MonitoringService.info('UI Manager initialized successfully', { operation: 'ui-init-complete' }, 'renderer');
            }
        } catch (error) {
            if (window.MonitoringService) {
                window.MonitoringService.error('Failed to initialize UI Manager', {
                    error: error.message,
                    stack: error.stack,
                    operation: 'ui-init'
                }, 'renderer');
            }
            throw error;
        }
    }

    /**
     * Setup global event listeners
     */
    setupEventListeners() {
        // Handle window resize
        window.addEventListener('resize', this.handleResize.bind(this));
        
        // Handle theme changes
        if (window.matchMedia) {
            const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
            mediaQuery.addEventListener('change', this.handleThemeChange.bind(this));
        }
    }

    /**
     * Handle window resize events
     */
    handleResize() {
        // Emit resize event for components that need to respond
        this.emitEvent('ui:resize', {
            width: window.innerWidth,
            height: window.innerHeight
        });
    }

    /**
     * Handle theme change events
     */
    handleThemeChange(event) {
        const isDark = event.matches;
        if (window.MonitoringService) {
            window.MonitoringService.info('Theme change detected', { isDark, operation: 'theme-change' }, 'renderer');
        }
        this.emitEvent('ui:theme-change', { isDark });
    }

    /**
     * Register a callback for specific events
     * @param {string} event - Event name
     * @param {Function} callback - Callback function
     */
    onEvent(event, callback) {
        if (!this.renderCallbacks.has(event)) {
            this.renderCallbacks.set(event, []);
        }
        this.renderCallbacks.get(event).push(callback);
    }

    /**
     * Emit an event to registered callbacks
     * @param {string} event - Event name
     * @param {any} data - Event data
     */
    emitEvent(event, data) {
        const callbacks = this.renderCallbacks.get(event) || [];
        callbacks.forEach(callback => {
            try {
                callback(data);
            } catch (error) {
                if (window.MonitoringService) {
                    window.MonitoringService.error('Error in UI event callback', {
                        event,
                        error: error.message,
                        operation: 'ui-event-callback'
                    }, 'renderer');
                }
            }
        });
    }

    /**
     * Create safe HTML element from template
     * @param {string} tag - HTML tag name
     * @param {Object} attributes - Element attributes
     * @param {string} content - Element content (will be sanitized)
     * @returns {HTMLElement} Created element
     */
    createElement(tag, attributes = {}, content = '') {
        try {
            const element = document.createElement(tag);
            
            // Set attributes safely
            Object.entries(attributes).forEach(([key, value]) => {
                if (key === 'className') {
                    element.className = DOMSafe?.sanitizeClassName(value) || '';
                } else if (key === 'href' && tag === 'a') {
                    element.href = DOMSafe?.sanitizeURL(value) || '#';
                } else if (typeof value === 'string') {
                    element.setAttribute(key, value);
                }
            });
            
            // Set content safely
            if (content) {
                if (DOMSafe) {
                    DOMSafe.setText(element, content);
                } else {
                    element.textContent = content;
                }
            }
            
            return element;
        } catch (error) {
            if (window.MonitoringService) {
                window.MonitoringService.error('Error creating element', {
                    tag,
                    attributes,
                    error: error.message,
                    operation: 'create-element'
                }, 'renderer');
            }
            throw error;
        }
    }

    /**
     * Show loading state on element
     * @param {HTMLElement} element - Element to show loading on
     * @param {string} message - Loading message
     */
    showLoading(element, message = 'Loading...') {
        if (!element) return;
        
        const loadingDiv = this.createElement('div', {
            className: 'loading-overlay'
        });
        
        const spinner = this.createElement('div', {
            className: 'loading-spinner'
        });
        
        const messageDiv = this.createElement('div', {
            className: 'loading-message'
        }, message);
        
        loadingDiv.appendChild(spinner);
        loadingDiv.appendChild(messageDiv);
        
        // Add loading styles if not present
        this.ensureLoadingStyles();
        
        element.classList.add('relative');
        element.appendChild(loadingDiv);
    }

    /**
     * Hide loading state from element
     * @param {HTMLElement} element - Element to hide loading from
     */
    hideLoading(element) {
        if (!element) return;
        
        const loadingOverlay = element.querySelector('.loading-overlay');
        if (loadingOverlay) {
            loadingOverlay.remove();
        }
    }

    /**
     * Show error state on element
     * @param {HTMLElement} element - Element to show error on
     * @param {string} message - Error message
     * @param {Function} retryCallback - Optional retry callback
     */
    showError(element, message = 'An error occurred', retryCallback = null) {
        if (!element) return;
        
        const errorDiv = this.createElement('div', {
            className: 'error-state'
        });
        
        const messageDiv = this.createElement('div', {
            className: 'error-message'
        }, message);
        
        errorDiv.appendChild(messageDiv);
        
        if (retryCallback) {
            const retryButton = this.createElement('button', {
                className: 'btn btn-primary btn-sm'
            }, 'Retry');
            
            retryButton.addEventListener('click', (e) => {
                e.preventDefault();
                this.hideError(element);
                retryCallback();
            });
            
            errorDiv.appendChild(retryButton);
        }
        
        // Add error styles if not present
        this.ensureErrorStyles();
        
        element.innerHTML = '';
        element.appendChild(errorDiv);
    }

    /**
     * Hide error state from element
     * @param {HTMLElement} element - Element to hide error from
     */
    hideError(element) {
        if (!element) return;
        
        const errorState = element.querySelector('.error-state');
        if (errorState) {
            errorState.remove();
        }
    }

    /**
     * Toggle element visibility with animation
     * @param {HTMLElement} element - Element to toggle
     * @param {boolean} show - Whether to show or hide
     */
    toggleVisibility(element, show) {
        if (!element) return;
        
        if (show) {
            element.classList.remove('hidden');
            element.classList.add('fade-in');
        } else {
            element.classList.add('fade-out');
            setTimeout(() => {
                element.classList.add('hidden');
                element.classList.remove('fade-out');
            }, 300);
        }
    }

    /**
     * Ensure loading styles are present
     */
    ensureLoadingStyles() {
        if (document.getElementById('ui-loading-styles')) return;
        
        const styles = `
            .loading-overlay {
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: rgba(255, 255, 255, 0.9);
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                z-index: 1000;
            }
            
            .loading-spinner {
                width: 32px;
                height: 32px;
                border: 3px solid var(--neutral-80);
                border-top: 3px solid var(--primary);
                border-radius: 50%;
                animation: spin 1s linear infinite;
            }
            
            .loading-message {
                margin-top: 12px;
                color: var(--text-secondary);
                font-size: 14px;
            }
            
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
            
            .fade-in {
                animation: fadeIn 0.3s ease-in;
            }
            
            .fade-out {
                animation: fadeOut 0.3s ease-out;
            }
            
            @keyframes fadeIn {
                from { opacity: 0; }
                to { opacity: 1; }
            }
            
            @keyframes fadeOut {
                from { opacity: 1; }
                to { opacity: 0; }
            }
        `;
        
        const styleSheet = document.createElement('style');
        styleSheet.id = 'ui-loading-styles';
        styleSheet.textContent = styles;
        document.head.appendChild(styleSheet);
    }

    /**
     * Ensure error styles are present
     */
    ensureErrorStyles() {
        if (document.getElementById('ui-error-styles')) return;
        
        const styles = `
            .error-state {
                padding: 20px;
                text-align: center;
                color: var(--error-color);
                background: var(--error-bg);
                border: 1px solid var(--error-color);
                border-radius: 8px;
                margin: 16px 0;
            }
            
            .error-message {
                margin-bottom: 12px;
                font-weight: 500;
            }
            
            .error-state .btn {
                margin-top: 8px;
            }
        `;
        
        const styleSheet = document.createElement('style');
        styleSheet.id = 'ui-error-styles';
        styleSheet.textContent = styles;
        document.head.appendChild(styleSheet);
    }

    /**
     * Start Microsoft Graph authentication flow using existing MSAL service
     */
    async startMicrosoftAuth() {
        try {
            UINotification.show('Redirecting to Microsoft login...', 'info');
            
            // Direct navigation to login endpoint (supports GET)
            window.location.href = '/api/auth/login';
            
        } catch (error) {
            console.error('Microsoft authentication failed:', error);
            UINotification.show(`Authentication failed: ${error.message}`, 'error');
        }
    }

    /**
     * Check authentication status using existing MSAL service
     */
    async refreshAuthenticationState() {
        try {
            // Check authentication status via existing status endpoint
            const response = await fetch('/api/status', { credentials: 'include' });
            
            if (!response.ok) {
                throw new Error('Failed to check authentication status');
            }

            const status = await response.json();
            
            // Check the actual response format from the API
            if (status.details?.msGraph?.authenticated) {
                this.showAuthenticatedState(status.details.msGraph);
            } else {
                this.showUnauthenticatedState();
            }
            
        } catch (error) {
            console.error('Failed to check authentication state:', error);
            this.showUnauthenticatedState();
        }
    }

    /**
     * Show authenticated state in UI
     */
    showAuthenticatedState(authInfo) {
        // Store user info for use in other UI elements
        this.currentUser = {
            name: authInfo.name || authInfo.user || 'User',
            email: authInfo.user || authInfo.email || '',
            sessionId: authInfo.sessionId
        };

        // Update login button to show authenticated state
        const loginButton = document.getElementById('device-auth-button');
        if (loginButton) {
            loginButton.textContent = '✅ Connected to Microsoft 365';
            loginButton.className = 'btn btn-success';
            loginButton.disabled = true;
        }

        // Show authentication status
        const authStatus = document.getElementById('auth-status');
        if (authStatus) {
            authStatus.classList.remove('hidden');
            const statusText = authStatus.querySelector('.status-text');
            if (statusText) {
                const userName = authInfo.name || authInfo.user || 'User';
                statusText.textContent = `Connected as ${userName}`;
            }
        }

        // Show user info if available
        if (authInfo.name || authInfo.user) {
            const displayName = authInfo.name || authInfo.user || 'User';
            UINotification.show(`Welcome ${displayName}!`, 'success');
        }

        // Show next step: Create session for MCP adapter
        this.showSessionCreationOption();
    }

    /**
     * Show session creation option after Microsoft authentication
     */
    showSessionCreationOption() {
        // Show the MCP token generation section
        const mcpTokenSection = document.getElementById('mcp-token-section');
        if (mcpTokenSection) {
            mcpTokenSection.classList.remove('hidden');

            // Set up event listeners for MCP token generation
            this.setupMcpTokenListeners();
        }

        // Show the external token section
        const externalTokenSection = document.getElementById('external-token-section');
        if (externalTokenSection) {
            externalTokenSection.classList.remove('hidden');

            // Set up event listeners for external token
            this.setupExternalTokenListeners();

            // Check for existing external token
            this.checkExternalTokenStatus();
        }

        // Hide the old adapter download section if it exists
        const downloadSection = document.getElementById('adapter-download-section');
        if (downloadSection) {
            downloadSection.classList.add('hidden');
        }
    }

    /**
     * Set up event listeners for MCP token generation
     */
    setupMcpTokenListeners() {
        // Generate MCP Token button
        const generateButton = document.getElementById('generate-mcp-token-button');
        if (generateButton) {
            generateButton.onclick = () => this.generateMcpToken();
        }

        // Copy token button
        const copyTokenButton = document.getElementById('copy-token-button');
        if (copyTokenButton) {
            copyTokenButton.onclick = () => this.copyToClipboard('mcp-token-text', 'Token copied to clipboard!');
        }

        // Copy config button
        const copyConfigButton = document.getElementById('copy-config-button');
        if (copyConfigButton) {
            copyConfigButton.onclick = () => this.copyToClipboard('config-example', 'Config copied to clipboard!');
        }

        // Download adapter button
        const downloadAdapterButton = document.getElementById('download-adapter-button');
        if (downloadAdapterButton) {
            downloadAdapterButton.onclick = () => this.downloadMcpAdapter();
        }
    }

    /**
     * Generate MCP bearer token
     */
    async generateMcpToken() {
        const generateButton = document.getElementById('generate-mcp-token-button');
        const resultDiv = document.getElementById('mcp-token-result');
        
        try {
            // Show loading state
            if (generateButton) {
                generateButton.disabled = true;
                generateButton.textContent = 'Generating...';
            }

            UINotification.show('Generating MCP token...', 'info');

            // Call the API to generate the token - using web auth endpoint
            const response = await fetch('/api/auth/generate-mcp-token', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include'
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error_description || 'Failed to generate token');
            }

            const tokenData = await response.json();

            // Display the token and config
            this.displayMcpToken(tokenData);

            UINotification.show('MCP token generated successfully!', 'success');

        } catch (error) {
            console.error('Error generating MCP token:', error);
            UINotification.show(`Failed to generate token: ${error.message}`, 'error');
        } finally {
            // Reset button state
            if (generateButton) {
                generateButton.disabled = false;
                generateButton.textContent = 'Generate MCP Token';
            }
        }
    }

    /**
     * Display the generated MCP token and configuration
     */
    displayMcpToken(tokenData) {
        const resultDiv = document.getElementById('mcp-token-result');
        const tokenText = document.getElementById('mcp-token-text');
        const configExample = document.getElementById('config-example');

        if (tokenText) {
            tokenText.value = tokenData.access_token;
        }

        if (configExample) {
            const serverUrl = window.location.origin;

            // SSE Transport config (recommended - no local files needed)
            const sseConfigJson = {
                mcpServers: {
                    "microsoft-365": {
                        transport: {
                            type: "sse",
                            url: `${serverUrl}/api/mcp/sse`,
                            headers: {
                                Authorization: `Bearer ${tokenData.access_token}`
                            }
                        }
                    }
                }
            };

            configExample.textContent = JSON.stringify(sseConfigJson, null, 2);
        }

        if (resultDiv) {
            resultDiv.classList.remove('hidden');
        }
    }

    /**
     * Download MCP adapter file
     */
    async downloadMcpAdapter() {
        try {
            UINotification.show('Downloading MCP adapter...', 'info');
            
            // Create a link element to trigger download
            const link = document.createElement('a');
            link.href = '/api/mcp-adapter.cjs';
            link.download = 'mcp-adapter.cjs';
            link.style.display = 'none';
            
            // Add to DOM, click, and remove
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            UINotification.show('MCP adapter downloaded successfully!', 'success');
            
        } catch (error) {
            console.error('Error downloading MCP adapter:', error);
            UINotification.show(`Failed to download adapter: ${error.message}`, 'error');
        }
    }

    /**
     * Copy text to clipboard
     */
    async copyToClipboard(elementId, successMessage) {
        const element = document.getElementById(elementId);
        if (!element) return;

        try {
            const text = element.value || element.textContent;
            await navigator.clipboard.writeText(text);
            UINotification.show(successMessage, 'success');
        } catch (error) {
            console.error('Failed to copy to clipboard:', error);
            
            // Fallback: select the text
            if (element.select) {
                element.select();
                element.setSelectionRange(0, 99999); // For mobile devices
            } else {
                const range = document.createRange();
                range.selectNode(element);
                window.getSelection().removeAllRanges();
                window.getSelection().addRange(range);
            }
            
            UINotification.show('Text selected - please copy manually (Ctrl+C)', 'warning');
        }
    }

    /**
     * Create MCP session for adapter generation
     */
    async createMCPSession() {
        try {
            UINotification.show('Creating MCP session...', 'info');
            
            // Step 1: Register device to get device code
            const registerResponse = await fetch('/api/auth/device/register', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include',
                body: JSON.stringify({
                    device_name: `MCP-Session-${Date.now()}`,
                    device_type: 'mcp-adapter'
                })
            });

            if (!registerResponse.ok) {
                const errorData = await registerResponse.json();
                throw new Error(errorData.error_description || 'Device registration failed');
            }

            const deviceData = await registerResponse.json();
            UINotification.show('Device registered, authorizing...', 'info');
            
            // Step 2: Authorize device with current user session
            // Since user is already authenticated, we can authorize the device directly
            const authResponse = await fetch('/api/auth/device/authorize', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include', // Include session cookies
                body: JSON.stringify({
                    user_code: deviceData.user_code, // Use the actual user_code from registration
                    user_id: 'current_session' // Use current session
                })
            });

            if (!authResponse.ok) {
                const errorData = await authResponse.json();
                throw new Error(errorData.error_description || 'Device authorization failed');
            }

            UINotification.show('MCP session created successfully!', 'success');
            
            // Show adapter download option with device data
            this.showAdapterDownloadOption({
                device_id: deviceData.device_code, // Use device_code as the session identifier
                device_code: deviceData.device_code,
                user_code: deviceData.user_code
            });
            
        } catch (error) {
            console.error('Session creation failed:', error);
            UINotification.show(`Session creation failed: ${error.message}`, 'error');
        }
    }

    /**
     * Show adapter download option after session creation
     */
    showAdapterDownloadOption(sessionData) {
        const downloadSection = document.getElementById('adapter-download-section');
        if (downloadSection) {
            // Store session data globally and in dataset for retrieval
            window.currentSessionData = sessionData;
            downloadSection.dataset.deviceId = sessionData.device_id;
            
            downloadSection.innerHTML = `
                <div class="adapter-download">
                    <h3>Session Ready</h3>
                    <p><strong>Device ID:</strong> ${sessionData.device_id}</p>
                    <p>Your session is ready! You can now test the API or generate an MCP token.</p>
                    <div class="adapter-buttons">
                        <button class="btn btn-info" onclick="UIManager.testAPIConnectivity('${sessionData.device_id}')">
                            🧪 Test API Connectivity
                        </button>
                        <button class="btn btn-primary" onclick="UIManager.generateMcpToken()">
                            🔑 Generate MCP Token
                        </button>
                        <button class="btn btn-secondary" onclick="UIManager.logout()">
                            Logout
                        </button>
                    </div>
                </div>
            `;
        }
    }

    /**
     * Test API connectivity for the device session
     */
    async testAPIConnectivity(deviceId) {
        try {
            UINotification.show('Testing MCP tools connectivity...', 'info');
            
            // Test the actual API endpoints that MCP tools call internally
            // These correspond to the MCP tools: getEvents, readMail, listFiles, findPeople
            const tests = [
                { 
                    name: 'Calendar Events (getEvents tool)', 
                    endpoint: '/api/v1/calendar',
                    method: 'GET',
                    params: '?limit=5'
                },
                { 
                    name: 'Recent Emails (readMail tool)', 
                    endpoint: '/api/v1/mail',
                    method: 'GET', 
                    params: '?limit=5'
                },
                { 
                    name: 'OneDrive Files (listFiles tool)', 
                    endpoint: '/api/v1/files',
                    method: 'GET',
                    params: '?limit=5'
                },
                { 
                    name: 'People Directory (findPeople tool)', 
                    endpoint: '/api/v1/people',
                    method: 'GET',
                    params: '?limit=5'
                }
            ];
            
            const results = [];
            
            for (const test of tests) {
                try {
                    const fullEndpoint = test.endpoint + (test.params || '');
                    const response = await fetch(fullEndpoint, {
                        method: test.method,
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        credentials: 'include'
                    });
                    
                    if (response.ok) {
                        const data = await response.json();
                        
                        // Extract meaningful data counts from different API responses
                        let itemCount = 'N/A';
                        if (Array.isArray(data.events)) {
                            itemCount = `${data.events.length} events`;
                        } else if (Array.isArray(data.messages)) {
                            itemCount = `${data.messages.length} messages`;
                        } else if (Array.isArray(data.files)) {
                            itemCount = `${data.files.length} files`;
                        } else if (Array.isArray(data.people)) {
                            itemCount = `${data.people.length} people`;
                        } else if (data.success !== undefined) {
                            itemCount = 'Tool executed successfully';
                        }
                        
                        results.push({ 
                            name: test.name, 
                            status: '✅ Success',
                            data: itemCount,
                            details: data 
                        });
                    } else {
                        results.push({ 
                            name: test.name, 
                            status: '❌ Failed', 
                            error: `${response.status} ${response.statusText}` 
                        });
                    }
                } catch (error) {
                    results.push({ 
                        name: test.name, 
                        status: '❌ Error', 
                        error: error.message 
                    });
                }
            }
            
            // Show results
            this.showAPITestResults(results);
            UINotification.show('MCP tools connectivity test completed!', 'success');
            
        } catch (error) {
            console.error('MCP tools connectivity test failed:', error);
            UINotification.show(`Test failed: ${error.message}`, 'error');
        }
    }

    /**
     * Show API test results
     */
    showAPITestResults(results) {
        const downloadSection = document.getElementById('adapter-download-section');
        if (downloadSection) {
            let resultsHtml = '';
            results.forEach((result) => {
                let html = `
                    <div class="test-result">
                        <h4>${result.name}</h4>
                        <p><strong>Status:</strong> ${result.status}</p>`;
                
                if (result.data) {
                    html += `
                        <pre><code>${result.data}</code></pre>`;
                }
                
                if (result.error) {
                    html += `
                        <p><strong>Error:</strong> ${result.error}</p>`;
                }
                
                if (result.details) {
                    html += `
                        <pre><code>${JSON.stringify(result.details, null, 2)}</code></pre>`;
                }
                
                html += `
                    </div>`;
                
                resultsHtml += html;
            });
            
            downloadSection.innerHTML = `
                <div class="adapter-download">
                    <h3>API Connectivity Test Results</h3>
                    <div class="test-results">
                        ${resultsHtml}
                    </div>
                    <div class="adapter-buttons">
                        <button class="btn btn-primary" onclick="UIManager.generateMcpToken()">
                            🔑 Generate MCP Token
                        </button>
                        <button class="btn btn-secondary" onclick="UIManager.logout()">
                            Logout
                        </button>
                    </div>
                </div>
            `;
        }
    }

    /**
     * Show unauthenticated state in UI
     */
    showUnauthenticatedState() {
        const loginButton = document.getElementById('device-auth-button');
        if (loginButton) {
            loginButton.textContent = 'Connect to Microsoft 365';
            loginButton.className = 'btn btn-primary';
            loginButton.disabled = false;
            loginButton.onclick = () => this.startMicrosoftAuth();
        }
    }

    /**
     * Register device using existing device auth system
     */
    async registerDevice() {
        try {
            UINotification.show('Registering device...', 'info');
            
            const response = await fetch('/api/auth/device/register', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include',
                body: JSON.stringify({
                    device_name: `Browser-${Date.now()}`,
                    device_type: 'browser'
                })
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error_description || 'Device registration failed');
            }

            const deviceData = await response.json();
            
            UINotification.show('Device registered successfully!', 'success');
            
            // Show device authorization flow
            this.showDeviceAuthFlow(deviceData);
            
        } catch (error) {
            console.error('Device registration failed:', error);
            UINotification.show(`Device registration failed: ${error.message}`, 'error');
        }
    }

    /**
     * Show device authorization flow
     */
    showDeviceAuthFlow(deviceData) {
        const downloadSection = document.getElementById('adapter-download-section');
        if (downloadSection) {
            downloadSection.innerHTML = `
                <div class="adapter-download">
                    <h3>Device Authorization</h3>
                    <p>Your device code: <strong>${deviceData.user_code}</strong></p>
                    <p>Visit: <a href="${deviceData.verification_uri}" target="_blank">${deviceData.verification_uri}</a></p>
                    <p>Enter the code above to authorize this device.</p>
                    <div class="adapter-buttons">
                        <button class="btn btn-primary" onclick="UIManager.pollForToken('${deviceData.device_id}')">
                            Check Authorization Status
                        </button>
                        <button class="btn btn-secondary" onclick="UIManager.copyUserCode('${deviceData.user_code}')">
                            Copy Code
                        </button>
                    </div>
                </div>
            `;
        }
    }

    /**
     * Poll for device token
     */
    async pollForToken(deviceId) {
        try {
            UINotification.show('Checking authorization status...', 'info');
            
            const response = await fetch('/api/auth/device/token', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include',
                body: JSON.stringify({
                    device_id: deviceId
                })
            });

            const result = await response.json();
            
            if (response.ok && result.access_token) {
                UINotification.show('Device authorized successfully!', 'success');
                this.showAdapterDownload(result);
            } else if (result.error === 'authorization_pending') {
                UINotification.show('Authorization pending. Please complete authorization in browser.', 'warning');
                // Continue polling
                setTimeout(() => this.pollForToken(deviceId), 5000);
            } else {
                throw new Error(result.error_description || 'Authorization failed');
            }
            
        } catch (error) {
            console.error('Token polling failed:', error);
            UINotification.show(`Authorization check failed: ${error.message}`, 'error');
        }
    }

    /**
     * Show token generation option after successful authorization
     */
    showAdapterDownload(tokenData) {
        const downloadSection = document.getElementById('adapter-download-section');
        if (downloadSection) {
            downloadSection.innerHTML = `
                <div class="adapter-download">
                    <h3>Authorization Complete</h3>
                    <p>Your device is authorized! You can now generate an MCP token.</p>
                    <div class="adapter-buttons">
                        <button class="btn btn-primary" onclick="UIManager.generateMcpToken()">
                            🔑 Generate MCP Token
                        </button>
                        <button class="btn btn-secondary" onclick="UIManager.logout()">
                            Logout
                        </button>
                    </div>
                </div>
            `;
        }
    }

    /**
     * Copy user code to clipboard
     */
    copyUserCode(userCode) {
        navigator.clipboard.writeText(userCode).then(() => {
            UINotification.show('User code copied to clipboard', 'success');
        }).catch(() => {
            UINotification.show('Failed to copy code', 'error');
        });
    }

    /**
     * Logout using existing MSAL service
     */
    async logout() {
        try {
            const response = await fetch('/api/auth/logout', {
                method: 'POST',
                credentials: 'include'
            });
            
            if (response.ok) {
                UINotification.show('Logged out successfully', 'success');
                this.showUnauthenticatedState();
                // Clear adapter download section
                const downloadSection = document.getElementById('adapter-download-section');
                if (downloadSection) {
                    downloadSection.classList.add('hidden');
                }
            } else {
                throw new Error('Logout failed');
            }
        } catch (error) {
            UINotification.show(`Logout failed: ${error.message}`, 'error');
        }
    }

    /**
     * Get current session data from window or UI state
     */
    getCurrentSessionData() {
        return window.currentSessionData || null;
    }

    // =============================================
    // External Token Management
    // =============================================

    /**
     * External token countdown interval
     */
    externalTokenCountdownInterval = null;

    /**
     * Setup event listeners for external token section
     */
    setupExternalTokenListeners() {
        // Inject token button
        const injectButton = document.getElementById('inject-external-token-button');
        if (injectButton) {
            injectButton.onclick = () => this.injectExternalToken();
        }

        // Clear token button
        const clearButton = document.getElementById('clear-external-token-button');
        if (clearButton) {
            clearButton.onclick = () => this.clearExternalToken();
        }

        // Toggle switch for using external token
        const toggle = document.getElementById('use-external-token-toggle');
        if (toggle) {
            toggle.onchange = (e) => this.toggleExternalTokenSource(e.target.checked);
        }
    }

    /**
     * Inject external token from input field
     * Uses login endpoint if not authenticated, inject endpoint if already authenticated
     */
    async injectExternalToken() {
        const tokenInput = document.getElementById('external-token-input');
        const injectButton = document.getElementById('inject-external-token-button');

        if (!tokenInput || !tokenInput.value.trim()) {
            UINotification.show('Please paste a token first', 'warning');
            return;
        }

        try {
            // Show loading state
            if (injectButton) {
                injectButton.disabled = true;
                injectButton.textContent = 'Validating...';
            }

            UINotification.show('Validating external token...', 'info');

            // Try the login endpoint first (works without auth)
            // This will create a session if successful
            const response = await fetch('/api/auth/external-token/login', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include',
                body: JSON.stringify({
                    access_token: tokenInput.value.trim()
                })
            });

            const result = await response.json();

            if (!response.ok) {
                throw new Error(result.message || result.error || 'Token validation failed');
            }

            // Clear input and show success
            tokenInput.value = '';

            if (result.authenticated) {
                UINotification.show(`Logged in as ${result.user?.name || result.user?.email}!`, 'success');

                // Update authenticated state
                this.showAuthenticatedState({
                    user: result.user?.email,
                    name: result.user?.name,
                    authenticated: true
                });
            } else {
                UINotification.show('External token injected successfully!', 'success');
            }

            // Display token info
            this.displayExternalTokenInfo(result.metadata);

        } catch (error) {
            console.error('External token injection failed:', error);
            UINotification.show(`Token injection failed: ${error.message}`, 'error');
        } finally {
            // Reset button state
            if (injectButton) {
                injectButton.disabled = false;
                injectButton.textContent = 'Login with Token';
            }
        }
    }

    /**
     * Display external token information
     */
    displayExternalTokenInfo(metadata) {
        const statusDiv = document.getElementById('external-token-status');
        const userSpan = document.getElementById('ext-token-user');
        const emailSpan = document.getElementById('ext-token-email');
        const scopeCountSpan = document.getElementById('ext-token-scope-count');
        const expiresSpan = document.getElementById('ext-token-expires');
        const scopesList = document.getElementById('ext-token-scopes-list');
        const toggle = document.getElementById('use-external-token-toggle');
        const warning = document.getElementById('ext-token-warning');

        if (userSpan) {
            userSpan.textContent = metadata.user?.name || 'Unknown';
        }

        if (emailSpan) {
            emailSpan.textContent = metadata.user?.email || 'Unknown';
        }

        if (scopeCountSpan) {
            scopeCountSpan.textContent = metadata.scopes?.length || 0;
        }

        if (expiresSpan) {
            const expiresAt = new Date(metadata.expires_at);
            expiresSpan.textContent = expiresAt.toLocaleString();
        }

        if (scopesList) {
            scopesList.innerHTML = metadata.scopes?.map(scope =>
                `<span class="scope-badge">${scope}</span>`
            ).join('') || 'No scopes';
        }

        if (toggle) {
            toggle.checked = true; // Default to using external token
        }

        if (warning) {
            warning.classList.toggle('hidden', !metadata.is_expiring_soon);
        }

        if (statusDiv) {
            statusDiv.classList.remove('hidden');
        }

        // Start countdown timer
        this.startExternalTokenCountdown(metadata.expires_at);
    }

    /**
     * Start countdown timer for external token expiry
     */
    startExternalTokenCountdown(expiresAt) {
        // Clear existing interval
        if (this.externalTokenCountdownInterval) {
            clearInterval(this.externalTokenCountdownInterval);
        }

        const countdownSpan = document.getElementById('ext-token-countdown');
        const warning = document.getElementById('ext-token-warning');
        const expiryTime = new Date(expiresAt).getTime();

        const updateCountdown = () => {
            const now = Date.now();
            const remaining = Math.max(0, expiryTime - now);

            if (remaining === 0) {
                if (countdownSpan) {
                    countdownSpan.textContent = 'EXPIRED';
                    countdownSpan.classList.add('expired');
                }
                if (warning) {
                    warning.classList.remove('hidden');
                    warning.textContent = 'Token has expired! Please paste a new token.';
                }
                clearInterval(this.externalTokenCountdownInterval);
                return;
            }

            const minutes = Math.floor(remaining / 60000);
            const seconds = Math.floor((remaining % 60000) / 1000);

            if (countdownSpan) {
                countdownSpan.textContent = `${minutes}m ${seconds}s`;
                countdownSpan.classList.toggle('warning', minutes < 10);
            }

            if (warning && minutes < 10) {
                warning.classList.remove('hidden');
            }
        };

        // Update immediately and then every second
        updateCountdown();
        this.externalTokenCountdownInterval = setInterval(updateCountdown, 1000);
    }

    /**
     * Clear external token
     */
    async clearExternalToken() {
        try {
            UINotification.show('Clearing external token...', 'info');

            const response = await fetch('/api/auth/external-token', {
                method: 'DELETE',
                credentials: 'include'
            });

            if (!response.ok) {
                const result = await response.json();
                throw new Error(result.message || 'Failed to clear token');
            }

            // Clear countdown interval
            if (this.externalTokenCountdownInterval) {
                clearInterval(this.externalTokenCountdownInterval);
                this.externalTokenCountdownInterval = null;
            }

            // Hide status section
            const statusDiv = document.getElementById('external-token-status');
            if (statusDiv) {
                statusDiv.classList.add('hidden');
            }

            // Reset toggle
            const toggle = document.getElementById('use-external-token-toggle');
            if (toggle) {
                toggle.checked = false;
            }

            UINotification.show('External token cleared', 'success');

        } catch (error) {
            console.error('Failed to clear external token:', error);
            UINotification.show(`Failed to clear token: ${error.message}`, 'error');
        }
    }

    /**
     * Toggle between external token and OAuth
     */
    async toggleExternalTokenSource(useExternal) {
        try {
            const response = await fetch('/api/auth/external-token/switch', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include',
                body: JSON.stringify({
                    source: useExternal ? 'external' : 'oauth'
                })
            });

            const result = await response.json();

            if (!response.ok) {
                throw new Error(result.message || 'Failed to switch token source');
            }

            UINotification.show(
                `Now using ${useExternal ? 'external enterprise' : 'OAuth'} token`,
                'success'
            );

        } catch (error) {
            console.error('Failed to switch token source:', error);
            UINotification.show(`Failed to switch: ${error.message}`, 'error');

            // Reset toggle on error
            const toggle = document.getElementById('use-external-token-toggle');
            if (toggle) {
                toggle.checked = !useExternal;
            }
        }
    }

    /**
     * Check and display external token status on page load
     */
    async checkExternalTokenStatus() {
        try {
            const response = await fetch('/api/auth/external-token/status', {
                credentials: 'include'
            });

            if (!response.ok) {
                return; // Not authenticated or no token
            }

            const status = await response.json();

            if (status.has_external_token && status.metadata) {
                this.displayExternalTokenInfo(status.metadata);

                // Set toggle based on current source
                const toggle = document.getElementById('use-external-token-toggle');
                if (toggle) {
                    toggle.checked = status.is_active;
                }
            }

        } catch (error) {
            // Silent fail - external token is optional
            console.debug('External token status check failed:', error.message);
        }
    }

    /**
     * Cleanup UI Manager resources
     */
    destroy() {
        // Clear external token countdown
        if (this.externalTokenCountdownInterval) {
            clearInterval(this.externalTokenCountdownInterval);
        }
        this.renderCallbacks.clear();
        this.initialized = false;
        if (window.MonitoringService) {
            window.MonitoringService.info('UI Manager destroyed', { operation: 'ui-cleanup' }, 'renderer');
        }
    }
}

// Make available globally
if (typeof window !== 'undefined') {
    // Create a global instance
    window.UIManagerInstance = new UIManager();
    
    // Make methods available as static methods for onclick handlers
    window.UIManager = {
        startMicrosoftAuth: () => window.UIManagerInstance.startMicrosoftAuth(),
        showAdapterDownloadOption: (sessionId) => window.UIManagerInstance.showAdapterDownloadOption(sessionId),
        refreshAuthenticationState: () => window.UIManagerInstance.refreshAuthenticationState(),
        showAuthenticatedState: (session) => window.UIManagerInstance.showAuthenticatedState(session),
        showUnauthenticatedState: () => window.UIManagerInstance.showUnauthenticatedState(),
        registerDevice: () => window.UIManagerInstance.registerDevice(),
        logout: () => window.UIManagerInstance.logout(),
        pollForToken: (deviceId) => window.UIManagerInstance.pollForToken(deviceId),
        copyUserCode: (userCode) => window.UIManagerInstance.copyUserCode(userCode),
        createMCPSession: () => window.UIManagerInstance.createMCPSession(),
        testAPIConnectivity: (deviceId) => window.UIManagerInstance.testAPIConnectivity(deviceId),
        // External token methods
        injectExternalToken: () => window.UIManagerInstance.injectExternalToken(),
        clearExternalToken: () => window.UIManagerInstance.clearExternalToken(),
        toggleExternalTokenSource: (useExternal) => window.UIManagerInstance.toggleExternalTokenSource(useExternal),
        checkExternalTokenStatus: () => window.UIManagerInstance.checkExternalTokenStatus()
    };
}