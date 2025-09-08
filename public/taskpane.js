/* global Office, OfficeExtension */

// TENS AI Office Add-in JavaScript
class TENSAIAddon {
    constructor() {
        this.apiKey = null;
        this.currentHost = null;
        this.isAuthenticated = false;
        this.currentResult = null;
        
        this.initializeOffice();
        this.bindEvents();
        this.loadStoredApiKey();
    }

    initializeOffice() {
        Office.onReady((info) => {
            this.currentHost = info.host;
            console.log(`TENS AI Add-in loaded in ${info.host}`);
            
            // Show appropriate UI based on host
            this.adaptUIForHost(info.host);
        });
    }

    adaptUIForHost(host) {
        const titleElement = document.querySelector('.ms-welcome__header h1');
        if (titleElement) {
            switch (host) {
                case Office.HostType.Excel:
                    titleElement.textContent = 'TENS AI - Excel Assistant';
                    break;
                case Office.HostType.Word:
                    titleElement.textContent = 'TENS AI - Word Assistant';
                    break;
                case Office.HostType.PowerPoint:
                    titleElement.textContent = 'TENS AI - PowerPoint Assistant';
                    break;
                case Office.HostType.Outlook:
                    titleElement.textContent = 'TENS AI - Outlook Assistant';
                    break;
            }
        }
    }

    bindEvents() {
        // Authentication
        document.getElementById('authenticate-btn').addEventListener('click', () => this.authenticate());
        document.getElementById('change-api-key-btn').addEventListener('click', () => this.showAuthSection());

        // Quick actions
        document.getElementById('analyze-text-btn').addEventListener('click', () => this.setQuickAction('analyze'));
        document.getElementById('generate-content-btn').addEventListener('click', () => this.setQuickAction('generate'));
        document.getElementById('improve-writing-btn').addEventListener('click', () => this.setQuickAction('improve'));
        document.getElementById('summarize-btn').addEventListener('click', () => this.setQuickAction('summarize'));

        // Content management
        document.getElementById('get-selection-btn').addEventListener('click', () => this.getSelectedText());
        document.getElementById('clear-content-btn').addEventListener('click', () => this.clearContent());

        // Prompt suggestions
        document.querySelectorAll('.suggestion-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const prompt = e.target.getAttribute('data-prompt');
                document.getElementById('ai-prompt').value = prompt;
                this.updateProcessButton();
            });
        });

        // Process button
        document.getElementById('process-btn').addEventListener('click', () => this.processWithAI());

        // Results
        document.getElementById('insert-result-btn').addEventListener('click', () => this.insertResult());
        document.getElementById('copy-result-btn').addEventListener('click', () => this.copyResult());

        // Tab switching
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => this.switchTab(e.target.getAttribute('data-tab')));
        });

        // Input validation
        document.getElementById('content-input').addEventListener('input', () => this.updateProcessButton());
        document.getElementById('ai-prompt').addEventListener('input', () => this.updateProcessButton());
    }

    loadStoredApiKey() {
        // Try to load API key from localStorage
        const storedKey = localStorage.getItem('tens_ai_api_key');
        if (storedKey) {
            this.apiKey = storedKey;
            this.isAuthenticated = true;
            this.showMainInterface();
            this.updateApiKeyDisplay();
        }
    }

    async authenticate() {
        const apiKeyInput = document.getElementById('api-key');
        const apiKey = apiKeyInput.value.trim();

        if (!apiKey) {
            this.showError('Please enter your API key');
            return;
        }

        this.showLoading('Validating API key...');

        try {
            // Test the API key by making a simple request
            const response = await fetch('/api/aimlgyan/health', {
                method: 'GET',
                headers: {
                    'x-api-key': apiKey
                }
            });

            if (response.ok) {
                this.apiKey = apiKey;
                this.isAuthenticated = true;
                localStorage.setItem('tens_ai_api_key', apiKey);
                this.showMainInterface();
                this.updateApiKeyDisplay();
                this.showSuccess('Successfully connected to TENS AI!');
            } else {
                throw new Error('Invalid API key');
            }
        } catch (error) {
            this.showError('Failed to authenticate. Please check your API key.');
            console.error('Authentication error:', error);
        } finally {
            this.hideLoading();
        }
    }

    showAuthSection() {
        document.getElementById('auth-section').style.display = 'block';
        document.getElementById('main-interface').style.display = 'none';
        document.getElementById('api-key').value = '';
    }

    showMainInterface() {
        document.getElementById('auth-section').style.display = 'none';
        document.getElementById('main-interface').style.display = 'block';
    }

    updateApiKeyDisplay() {
        if (this.apiKey) {
            const display = this.apiKey.substring(0, 8) + '••••••••••••••••';
            document.getElementById('api-key-display').textContent = display;
        }
    }

    setQuickAction(action) {
        const promptInput = document.getElementById('ai-prompt');
        
        switch (action) {
            case 'analyze':
                promptInput.value = 'Analyze this text for sentiment, key themes, and provide insights';
                break;
            case 'generate':
                promptInput.value = 'Generate relevant content based on this context';
                break;
            case 'improve':
                promptInput.value = 'Improve the grammar, clarity, and overall quality of this text';
                break;
            case 'summarize':
                promptInput.value = 'Provide a concise summary of the key points in this text';
                break;
        }
        
        this.updateProcessButton();
    }

    async getSelectedText() {
        try {
            let selectedText = '';

            switch (this.currentHost) {
                case Office.HostType.Word:
                    selectedText = await this.getWordSelection();
                    break;
                case Office.HostType.Excel:
                    selectedText = await this.getExcelSelection();
                    break;
                case Office.HostType.PowerPoint:
                    selectedText = await this.getPowerPointSelection();
                    break;
                case Office.HostType.Outlook:
                    selectedText = await this.getOutlookSelection();
                    break;
            }

            if (selectedText) {
                document.getElementById('content-input').value = selectedText;
                this.updateProcessButton();
                this.showSuccess('Selected text loaded successfully');
            } else {
                this.showError('No text selected. Please select some text in your document.');
            }
        } catch (error) {
            this.showError('Failed to get selected text: ' + error.message);
            console.error('Get selection error:', error);
        }
    }

    async getWordSelection() {
        return new Promise((resolve, reject) => {
            Word.run(async (context) => {
                try {
                    const selection = context.document.getSelection();
                    selection.load('text');
                    await context.sync();
                    resolve(selection.text);
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    async getExcelSelection() {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                try {
                    const range = context.workbook.getSelectedRange();
                    range.load('text');
                    await context.sync();
                    resolve(range.text);
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    async getPowerPointSelection() {
        return new Promise((resolve, reject) => {
            PowerPoint.run(async (context) => {
                try {
                    const selection = context.presentation.getSelectedShapes();
                    selection.load('items');
                    await context.sync();
                    
                    let text = '';
                    selection.items.forEach(shape => {
                        if (shape.textFrame && shape.textFrame.textRange) {
                            text += shape.textFrame.textRange.text + '\n';
                        }
                    });
                    resolve(text.trim());
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    async getOutlookSelection() {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(new Error(result.error.message));
                }
            });
        });
    }

    clearContent() {
        document.getElementById('content-input').value = '';
        document.getElementById('ai-prompt').value = '';
        document.getElementById('results-section').style.display = 'none';
        this.updateProcessButton();
    }

    updateProcessButton() {
        const content = document.getElementById('content-input').value.trim();
        const prompt = document.getElementById('ai-prompt').value.trim();
        const processBtn = document.getElementById('process-btn');
        
        processBtn.disabled = !this.isAuthenticated || !content || !prompt;
    }

    async processWithAI() {
        if (!this.isAuthenticated) {
            this.showError('Please authenticate first');
            return;
        }

        const content = document.getElementById('content-input').value.trim();
        const prompt = document.getElementById('ai-prompt').value.trim();

        if (!content || !prompt) {
            this.showError('Please provide both content and prompt');
            return;
        }

        this.showLoading('Processing with TENS AI...');

        try {
            const response = await fetch('/api/generate-content', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'x-api-key': this.apiKey
                },
                body: JSON.stringify({
                    prompt: prompt,
                    contentType: 'text',
                    context: content
                })
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const result = await response.json();
            this.currentResult = result;
            this.displayResults(result);
            this.showSuccess('AI processing completed successfully!');

        } catch (error) {
            this.showError('Failed to process with AI: ' + error.message);
            console.error('AI processing error:', error);
        } finally {
            this.hideLoading();
        }
    }

    displayResults(result) {
        const resultsSection = document.getElementById('results-section');
        const resultContent = document.getElementById('result-content');
        const metadataContent = document.getElementById('metadata-content');

        // Display main result
        resultContent.textContent = result.content || result.text || 'No result content available';

        // Display metadata
        const metadata = {
            'Model': result.model || 'Unknown',
            'Tokens Used': result.tokens || 'N/A',
            'Processing Time': result.processingTime || 'N/A',
            'Confidence': result.confidence || 'N/A',
            'Timestamp': new Date().toLocaleString()
        };

        metadataContent.innerHTML = Object.entries(metadata)
            .map(([key, value]) => `<div><strong>${key}:</strong> ${value}</div>`)
            .join('');

        resultsSection.style.display = 'block';
    }

    switchTab(tabName) {
        // Update tab buttons
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        // Update tab content
        document.querySelectorAll('.tab-pane').forEach(pane => {
            pane.classList.remove('active');
        });
        document.getElementById(`${tabName}-tab`).classList.add('active');
    }

    async insertResult() {
        if (!this.currentResult) {
            this.showError('No result to insert');
            return;
        }

        const resultText = this.currentResult.content || this.currentResult.text;
        if (!resultText) {
            this.showError('No result content to insert');
            return;
        }

        try {
            switch (this.currentHost) {
                case Office.HostType.Word:
                    await this.insertIntoWord(resultText);
                    break;
                case Office.HostType.Excel:
                    await this.insertIntoExcel(resultText);
                    break;
                case Office.HostType.PowerPoint:
                    await this.insertIntoPowerPoint(resultText);
                    break;
                case Office.HostType.Outlook:
                    await this.insertIntoOutlook(resultText);
                    break;
            }
            this.showSuccess('Result inserted successfully!');
        } catch (error) {
            this.showError('Failed to insert result: ' + error.message);
            console.error('Insert error:', error);
        }
    }

    async insertIntoWord(text) {
        return new Promise((resolve, reject) => {
            Word.run(async (context) => {
                try {
                    const selection = context.document.getSelection();
                    selection.insertText(text, Word.InsertLocation.replace);
                    await context.sync();
                    resolve();
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    async insertIntoExcel(text) {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                try {
                    const range = context.workbook.getSelectedRange();
                    range.values = [[text]];
                    await context.sync();
                    resolve();
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    async insertIntoPowerPoint(text) {
        return new Promise((resolve, reject) => {
            PowerPoint.run(async (context) => {
                try {
                    const selection = context.presentation.getSelectedShapes();
                    selection.load('items');
                    await context.sync();
                    
                    if (selection.items.length > 0) {
                        const shape = selection.items[0];
                        if (shape.textFrame) {
                            shape.textFrame.textRange.text = text;
                        }
                    }
                    await context.sync();
                    resolve();
                } catch (error) {
                    reject(error);
                }
            });
        });
    }

    async insertIntoOutlook(text) {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.item.body.setSelectedDataAsync(text, {
                coercionType: Office.CoercionType.Text
            }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                    reject(new Error(result.error.message));
                }
            });
        });
    }

    async copyResult() {
        if (!this.currentResult) {
            this.showError('No result to copy');
            return;
        }

        const resultText = this.currentResult.content || this.currentResult.text;
        if (!resultText) {
            this.showError('No result content to copy');
            return;
        }

        try {
            await navigator.clipboard.writeText(resultText);
            this.showSuccess('Result copied to clipboard!');
        } catch (error) {
            this.showError('Failed to copy to clipboard: ' + error.message);
            console.error('Copy error:', error);
        }
    }

    showLoading(message = 'Loading...') {
        const overlay = document.getElementById('loading-overlay');
        const messageElement = overlay.querySelector('p');
        messageElement.textContent = message;
        overlay.style.display = 'flex';
    }

    hideLoading() {
        document.getElementById('loading-overlay').style.display = 'none';
    }

    showError(message) {
        const errorElement = document.getElementById('error-message');
        const errorText = document.getElementById('error-text');
        errorText.textContent = message;
        errorElement.style.display = 'flex';
        
        // Auto-hide after 5 seconds
        setTimeout(() => {
            errorElement.style.display = 'none';
        }, 5000);
    }

    showSuccess(message) {
        const successElement = document.getElementById('success-message');
        const successText = document.getElementById('success-text');
        successText.textContent = message;
        successElement.style.display = 'flex';
        
        // Auto-hide after 3 seconds
        setTimeout(() => {
            successElement.style.display = 'none';
        }, 3000);
    }
}

// Initialize the add-in when the page loads
document.addEventListener('DOMContentLoaded', () => {
    new TENSAIAddon();
});
