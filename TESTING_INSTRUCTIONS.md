# TENS AI Office Add-in Testing Instructions
## Updated for www.dev2.tens-ai.com

## 🚀 Quick Start Testing (5 minutes)

### Step 1: Start the Server
```bash
cd "/Users/vinaykumar/Desktop/Outlook Addon"
npm install
npm start
```

### Step 2: Test in Browser
1. Open browser and go to: `https://localhost:3000/taskpane.html`
2. You should see the TENS AI interface
3. If you get SSL warning, click "Advanced" → "Proceed to localhost"

### Step 3: Test in Office
1. Open Microsoft Word/Excel/PowerPoint/Outlook
2. Go to **Insert** → **Add-ins** → **Upload My Add-in**
3. Upload the `manifest.xml` file
4. Look for "TENS AI" in the Home ribbon
5. Click "AI Assistant" to open the task pane

## 📋 Complete Testing Checklist

### ✅ Phase 1: Server Test (2 minutes)

**Step 1.1: Check Server Health**
```bash
# In terminal, run:
curl -k https://localhost:3000/health
```
Expected: `{"status":"healthy","timestamp":"...","version":"1.0.0"}`

**Step 1.2: Test Web Interface**
- Open browser: `https://localhost:3000/taskpane.html`
- Should see: "TENS AI Assistant" interface
- Check: No JavaScript errors in browser console (F12)

**Step 1.3: Test API Connection**
```bash
# Test connection to dev2.tens-ai.com
curl -k https://localhost:3000/api/aimlgyan/health
```
Expected: Connection status response

### ✅ Phase 2: Office Integration Test (5 minutes)

**Step 2.1: Install Add-in in Word**
1. Open Microsoft Word
2. Go to Insert → Add-ins → Upload My Add-in
3. Select the `manifest.xml` file from your project folder
4. Click Upload
5. Look for "TENS AI" group in Home ribbon
6. Click "AI Assistant" button
7. Task pane should open on the right

**Step 2.2: Test in Other Office Apps**
Repeat Step 2.1 for:
- Excel
- PowerPoint  
- Outlook

**Step 2.3: Verify Manifest URLs**
- Check that manifest.xml points to correct domains:
  - `https://www.dev2.tens-ai.com`
  - `https://api.dev2.tens-ai.com`

### ✅ Phase 3: Authentication Test (3 minutes)

**Step 3.1: Test API Key Input**
1. In the task pane, you'll see an API key field
2. Click "Get your API key" link - should open `https://www.dev2.tens-ai.com/api-keys`
3. Enter a test API key
4. Click "Connect"
5. Expected: Connection attempt to `https://api.dev2.tens-ai.com`

**Step 3.2: Test with Valid API Key**
1. Get API key from `https://www.dev2.tens-ai.com/api-keys`
2. Enter the real API key
3. Click "Connect"
4. Expected: Success message and main interface appears

**Step 3.3: Test API Key Storage**
1. Verify API key is stored securely in browser
2. Refresh the page - API key should persist
3. Check that API key is not exposed in network requests

### ✅ Phase 4: Text Selection Test (3 minutes)

**Step 4.1: Test in Word**
1. Type some text in Word document
2. Select a portion of the text
3. In add-in, click "Get Selected Text"
4. Expected: Selected text appears in content area

**Step 4.2: Test in Excel**
1. Enter text in Excel cells
2. Select a range of cells
3. In add-in, click "Get Selected Text"
4. Expected: Cell contents appear in content area

**Step 4.3: Test in PowerPoint**
1. Create a slide with text
2. Select text in a text box
3. In add-in, click "Get Selected Text"
4. Expected: Selected text appears in content area

**Step 4.4: Test in Outlook**
1. Create a new email
2. Type content in the email body
3. In add-in, click "Get Selected Text"
4. Expected: Email content appears in content area

### ✅ Phase 5: AI Processing Test (5 minutes)

**Step 5.1: Test Content Processing**
1. Enter some text in the content area (or use "Get Selected Text")
2. In the AI Prompt field, type: "Summarize this text"
3. Click "Process with AI"
4. Expected: 
   - Loading spinner appears
   - Request sent to `https://api.dev2.tens-ai.com`
   - Results show in Results section

**Step 5.2: Test Quick Actions**
1. Click "Analyze Text" button
2. Notice the prompt field auto-fills
3. Click "Process with AI"
4. Expected: Analysis results from dev2.tens-ai.com API
5. Repeat for other quick action buttons:
   - "Generate Content"
   - "Improve Writing"
   - "Summarize"

**Step 5.3: Test API Integration**
1. Monitor network requests (F12 → Network tab)
2. Verify requests go to `https://api.dev2.tens-ai.com`
3. Check API responses are handled correctly
4. Test error handling for API failures

### ✅ Phase 6: Results Management Test (2 minutes)

**Step 6.1: Test Result Display**
1. After processing, results should appear in "Results" section
2. Click "Result" and "Metadata" tabs
3. Expected: Content switches between tabs
4. Verify metadata shows API response details

**Step 6.2: Test Insert Function**
1. Click "Insert into Document"
2. Expected: Result text appears in your Office document
3. Test in different Office applications

**Step 6.3: Test Copy Function**
1. Click "Copy to Clipboard"
2. Paste in a text editor
3. Expected: Result text is copied correctly

### ✅ Phase 7: Settings Test (2 minutes)

**Step 7.1: Test API Key Management**
1. Click "Change" next to the API key display
2. Expected: Authentication section appears
3. Enter a new API key
4. Expected: Key is updated and stored

**Step 7.2: Test Response Length Settings**
1. Change the response length setting
2. Process some content
3. Expected: Response length matches the setting

**Step 7.3: Test AI Model Selection**
1. Change the AI model setting
2. Process some content
3. Expected: Selected model is used in API requests

### ✅ Phase 8: Error Handling Test (3 minutes)

**Step 8.1: Test Network Errors**
1. Disconnect from internet
2. Try to process content
3. Expected: Appropriate error message appears

**Step 8.2: Test API Errors**
1. Use an invalid API key
2. Try to process content
3. Expected: Error handling works correctly
4. Check error messages are user-friendly

**Step 8.3: Test Empty Input**
1. Leave content area empty
2. Try to process
3. Expected: Validation error appears

**Step 8.4: Test API Timeout**
1. Test with slow network connection
2. Expected: Timeout handling works correctly

### ✅ Phase 9: Cross-Application Test (5 minutes)

**Step 9.1: Test Word Integration**
- Create a document with various content types
- Test all add-in features
- Verify text insertion works correctly
- Test with different document formats

**Step 9.2: Test Excel Integration**
- Create a spreadsheet with data
- Test text selection from cells
- Verify content generation works
- Test with different cell formats

**Step 9.3: Test PowerPoint Integration**
- Create a presentation with slides
- Test text selection from text boxes
- Verify content insertion works
- Test with different slide layouts

**Step 9.4: Test Outlook Integration**
- Create emails with different content
- Test text selection from email body
- Verify content generation and insertion
- Test with different email formats

## 🔧 Troubleshooting

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
2. Check that `https://www.dev2.tens-ai.com` is accessible
3. Test the API directly: `curl -H "Authorization: Bearer YOUR_API_KEY" https://api.dev2.tens-ai.com/health`
4. Check network connectivity to dev2.tens-ai.com

#### Domain Configuration Issues
1. Verify manifest.xml contains correct dev2.tens-ai.com URLs
2. Check CORS configuration allows dev2.tens-ai.com
3. Ensure API endpoints point to api.dev2.tens-ai.com

### Debug Commands

```bash
# Check setup
node test-setup.js

# Test server health
curl -k https://localhost:3000/health

# Test API connection
curl -k https://localhost:3000/api/aimlgyan/health

# Check port usage
lsof -i :3000

# Test dev2.tens-ai.com connectivity
curl -I https://www.dev2.tens-ai.com
curl -I https://api.dev2.tens-ai.com
```

## 🎯 Testing Checklist Summary

### ✅ Server Tests
- [ ] Server starts without errors
- [ ] Health endpoint responds
- [ ] Web interface loads in browser
- [ ] No JavaScript errors in console
- [ ] API connection to dev2.tens-ai.com works

### ✅ Office Integration Tests
- [ ] Add-in installs in Word
- [ ] Add-in installs in Excel
- [ ] Add-in installs in PowerPoint
- [ ] Add-in installs in Outlook
- [ ] Task pane opens correctly
- [ ] Interface displays properly
- [ ] Manifest URLs point to dev2.tens-ai.com

### ✅ Authentication Tests
- [ ] Invalid API key shows error
- [ ] Valid API key shows success
- [ ] Main interface appears after auth
- [ ] API key is stored securely
- [ ] API key link points to dev2.tens-ai.com

### ✅ Functionality Tests
- [ ] Text selection works in Word
- [ ] Text selection works in Excel
- [ ] Text selection works in PowerPoint
- [ ] Text selection works in Outlook
- [ ] AI processing works (with valid API key)
- [ ] Results display correctly
- [ ] Insert function works
- [ ] Copy function works
- [ ] Quick actions work
- [ ] Settings can be changed
- [ ] API requests go to api.dev2.tens-ai.com

### ✅ Error Handling Tests
- [ ] Network errors handled gracefully
- [ ] Invalid inputs show appropriate errors
- [ ] Empty content shows validation error
- [ ] API errors display user-friendly messages
- [ ] Timeout handling works correctly

## 🚀 Final Verification

After completing all tests:
1. [ ] All core features work
2. [ ] No critical errors
3. [ ] Performance is acceptable
4. [ ] Security requirements met
5. [ ] User experience is smooth
6. [ ] All URLs point to dev2.tens-ai.com
7. [ ] API integration works correctly

## 📞 Support

For issues and support:
- Check the troubleshooting section above
- Review server logs for error details
- Test connectivity to dev2.tens-ai.com
- Contact support at https://www.dev2.tens-ai.com/support

## 🎉 Success Criteria

The add-in is working correctly when:
- ✅ Server starts and serves the interface
- ✅ Add-in loads in all Office applications
- ✅ Authentication works with dev2.tens-ai.com API
- ✅ Text selection works across all Office apps
- ✅ AI processing returns results from dev2.tens-ai.com
- ✅ Results can be inserted into documents
- ✅ Error handling works gracefully
- ✅ All URLs point to the correct dev2.tens-ai.com domains
