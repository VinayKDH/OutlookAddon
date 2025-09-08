# AI ML Gyan Office Add-in

A comprehensive Microsoft Office add-in that integrates AI ML Gyan's powerful AI capabilities directly into Word, Excel, PowerPoint, and Outlook.

## Features

- **Multi-Office Support**: Works seamlessly across Word, Excel, PowerPoint, and Outlook
- **AI-Powered Content**: Generate, analyze, and improve content using AI ML Gyan's API
- **Smart Text Processing**: Analyze sentiment, extract keywords, and summarize content
- **Content Generation**: Create professional content based on prompts and context
- **Writing Enhancement**: Improve grammar, clarity, and overall text quality
- **Easy Integration**: Simple API key authentication and intuitive interface

## Prerequisites

- Node.js (version 16 or higher)
- npm (comes with Node.js)
- Microsoft Office 365 or Office 2019/2021
- AI ML Gyan API key (get from https://www.dev2.tens-ai.com/api-keys)

## Installation & Setup

### Step 1: Install Dependencies

```bash
# Navigate to the project directory
cd "/Users/vinaykumar/Desktop/Outlook Addon"

# Install Node.js dependencies
npm install
```

### Step 2: Configure Environment

```bash
# Copy the environment template
cp env.example .env

# Edit the .env file with your configuration
# At minimum, you need to set your AI ML Gyan API key
```

### Step 3: Start the Development Server

```bash
# Start the server in development mode
npm run dev

# Or start in production mode
npm start
```

The server will start on `https://localhost:3000`

## Manual Testing Instructions

### Phase 1: Server Testing

#### 1.1 Test Server Health
```bash
# Open a new terminal and test the health endpoint
curl -k https://localhost:3000/health

# Expected response:
# {"status":"healthy","timestamp":"2024-01-01T00:00:00.000Z","version":"1.0.0"}
```

#### 1.2 Test Static Files
- Open browser and navigate to: `https://localhost:3000/taskpane.html`
- You should see the AI ML Gyan Assistant interface
- Check that all CSS and JavaScript files load without errors

#### 1.3 Test API Endpoints
```bash
# Test AI ML Gyan health check (without API key)
curl -k https://localhost:3000/api/aimlgyan/health

# Test with API key (replace YOUR_API_KEY)
curl -k -H "x-api-key: YOUR_API_KEY" https://localhost:3000/api/aimlgyan/health
```

### Phase 2: Office Integration Testing

#### 2.1 Install the Add-in in Office

**For Office 365 Online:**
1. Open any Office app (Word, Excel, PowerPoint, or Outlook)
2. Go to Insert > Add-ins > Upload My Add-in
3. Upload the `manifest.xml` file
4. The add-in should appear in the ribbon

**For Office Desktop:**
1. Open any Office app
2. Go to Insert > My Add-ins > Upload My Add-in
3. Browse to the `manifest.xml` file
4. Click Upload

**For Development Testing:**
1. Open Office app
2. Go to Insert > Add-ins > Upload My Add-in
3. Use the manifest URL: `https://localhost:3000/manifest.xml`
4. Click Upload

#### 2.2 Test Add-in Loading
1. After installation, look for "AI ML Gyan" group in the Home ribbon
2. Click the "AI Assistant" button
3. The task pane should open on the right side
4. Verify the interface loads correctly

### Phase 3: Authentication Testing

#### 3.1 Test API Key Authentication
1. In the task pane, enter your AI ML Gyan API key
2. Click "Connect"
3. Verify success message appears
4. Check that the main interface becomes visible

#### 3.2 Test Invalid API Key
1. Enter an invalid API key
2. Click "Connect"
3. Verify error message appears
4. Check that authentication fails gracefully

### Phase 4: Core Functionality Testing

#### 4.1 Test Text Selection (Word)
1. Open Microsoft Word
2. Type some sample text
3. Select a portion of the text
4. In the add-in, click "Get Selected Text"
5. Verify the selected text appears in the content area

#### 4.2 Test Text Selection (Excel)
1. Open Microsoft Excel
2. Enter some text in cells
3. Select a range of cells
4. In the add-in, click "Get Selected Text"
5. Verify the cell contents appear in the content area

#### 4.3 Test Text Selection (PowerPoint)
1. Open Microsoft PowerPoint
2. Create a slide with text
3. Select text in a text box
4. In the add-in, click "Get Selected Text"
5. Verify the selected text appears in the content area

#### 4.4 Test Text Selection (Outlook)
1. Open Microsoft Outlook
2. Create a new email
3. Type some content in the email body
4. In the add-in, click "Get Selected Text"
5. Verify the email content appears in the content area

### Phase 5: AI Processing Testing

#### 5.1 Test Content Analysis
1. Enter some text in the content area
2. Set prompt to "Analyze this text for sentiment and key themes"
3. Click "Process with AI"
4. Verify results appear in the results section
5. Check that metadata is displayed

#### 5.2 Test Content Generation
1. Enter context text in the content area
2. Set prompt to "Generate a professional summary of this content"
3. Click "Process with AI"
4. Verify generated content appears
5. Test the "Insert into Document" functionality

#### 5.3 Test Writing Improvement
1. Enter text with grammar issues
2. Set prompt to "Improve the grammar and clarity of this text"
3. Click "Process with AI"
4. Verify improved text is generated
5. Test inserting the improved text

#### 5.4 Test Quick Actions
1. Click "Analyze Text" quick action button
2. Verify prompt is automatically filled
3. Click "Process with AI"
4. Repeat for other quick actions (Generate Content, Improve Writing, Summarize)

### Phase 6: Results Management Testing

#### 6.1 Test Result Insertion
1. Process some content with AI
2. Click "Insert into Document"
3. Verify the result is inserted into the Office document
4. Test in different Office applications

#### 6.2 Test Copy to Clipboard
1. Process some content with AI
2. Click "Copy to Clipboard"
3. Paste into a text editor
4. Verify the content was copied correctly

#### 6.3 Test Tab Switching
1. Process content with AI
2. Click between "Result" and "Metadata" tabs
3. Verify content switches correctly

### Phase 7: Settings Testing

#### 7.1 Test API Key Management
1. Click "Change" next to the API key display
2. Verify authentication section appears
3. Enter a new API key
4. Verify the key is updated

#### 7.2 Test Response Length Settings
1. Change the response length setting
2. Process some content
3. Verify the response length matches the setting

#### 7.3 Test AI Model Selection
1. Change the AI model setting
2. Process some content
3. Verify the selected model is used

### Phase 8: Error Handling Testing

#### 8.1 Test Network Errors
1. Disconnect from internet
2. Try to process content
3. Verify appropriate error message appears

#### 8.2 Test API Errors
1. Use an invalid API key
2. Try to process content
3. Verify error handling works correctly

#### 8.3 Test Empty Input
1. Leave content area empty
2. Try to process
3. Verify validation error appears

### Phase 9: Cross-Application Testing

#### 9.1 Test Word Integration
- Create a document with various content types
- Test all add-in features
- Verify text insertion works correctly

#### 9.2 Test Excel Integration
- Create a spreadsheet with data
- Test text selection from cells
- Verify content generation works

#### 9.3 Test PowerPoint Integration
- Create a presentation with slides
- Test text selection from text boxes
- Verify content insertion works

#### 9.4 Test Outlook Integration
- Create emails with different content
- Test text selection from email body
- Verify content generation and insertion

## Troubleshooting

### Common Issues

#### Server Won't Start
```bash
# Check if port 3000 is in use
lsof -i :3000

# Kill process if needed
kill -9 <PID>

# Try a different port
PORT=3001 npm start
```

#### SSL Certificate Issues
```bash
# For development, you can use self-signed certificates
# The browser will show a security warning - click "Advanced" and "Proceed"
```

#### Office Add-in Won't Load
1. Check that the manifest.xml is accessible at `https://localhost:3000/manifest.xml`
2. Verify all URLs in the manifest use HTTPS
3. Check browser console for JavaScript errors
4. Ensure Office app is up to date

#### API Connection Issues
1. Verify your API key is correct
2. Check that https://www.dev2.tens-ai.com is accessible
3. Test the API directly: `curl -H "Authorization: Bearer YOUR_API_KEY" https://api.dev2.tens-ai.com/health`

### Debug Mode

Enable debug logging:
```bash
# Set debug environment variable
DEBUG=* npm start

# Or check server logs
tail -f logs/app.log
```

## Production Deployment

### Using Docker
```bash
# Build the Docker image
docker build -t aimlgyan-office-addon .

# Run with Docker Compose
docker-compose up -d
```

### Manual Deployment
1. Set up a web server (nginx, Apache)
2. Configure SSL certificates
3. Update manifest.xml with production URLs
4. Deploy the application files
5. Configure reverse proxy for API calls

## Support

For issues and support:
- Check the troubleshooting section above
- Review server logs for error details
- Contact AI ML Gyan support at https://www.dev2.tens-ai.com/support

## License

MIT License - see LICENSE file for details.

# OutlookAddon
