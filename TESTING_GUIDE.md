# Manual Testing Guide for TENS AI Office Add-in

## Quick Start Testing (5 minutes)

### 1. Start the Server
```bash
cd "/Users/vinaykumar/Desktop/Outlook Addon"
npm install
npm start
```

### 2. Test in Browser
1. Open browser and go to: `https://localhost:3000/taskpane.html`
2. You should see the TENS AI interface
3. If you get SSL warning, click "Advanced" → "Proceed to localhost"

### 3. Test in Office
1. Open Microsoft Word/Excel/PowerPoint/Outlook
2. Go to Insert → Add-ins → Upload My Add-in
3. Upload the `manifest.xml` file
4. Look for "TENS AI" in the Home ribbon
5. Click "AI Assistant" to open the task pane

## Detailed Testing Steps

### Phase 1: Basic Server Test (2 minutes)

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

### Phase 2: Office Integration Test (5 minutes)

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

### Phase 3: Authentication Test (2 minutes)

**Step 3.1: Test API Key Input**
1. In the task pane, you'll see an API key field
2. Enter any text (we'll test with a dummy key first)
3. Click "Connect"
4. Expected: Error message (since it's not a real API key)

**Step 3.2: Test with Real API Key (if you have one)**
1. Get API key from https://www.dev2.tens-ai.com/api-keys
2. Enter the real API key
3. Click "Connect"
4. Expected: Success message and main interface appears

### Phase 4: Text Selection Test (3 minutes)

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

### Phase 5: AI Processing Test (5 minutes)

**Step 5.1: Test Content Processing**
1. Enter some text in the content area (or use "Get Selected Text")
2. In the AI Prompt field, type: "Summarize this text"
3. Click "Process with AI"
4. Expected: Loading spinner appears, then results show

**Step 5.2: Test Quick Actions**
1. Click "Analyze Text" button
2. Notice the prompt field auto-fills
3. Click "Process with AI"
4. Repeat for other quick action buttons

### Phase 6: Results Test (2 minutes)

**Step 6.1: Test Result Display**
1. After processing, results should appear in "Results" section
2. Click "Result" and "Metadata" tabs
3. Expected: Content switches between tabs

**Step 6.2: Test Insert Function**
1. Click "Insert into Document"
2. Expected: Result text appears in your Office document

**Step 6.3: Test Copy Function**
1. Click "Copy to Clipboard"
2. Paste in a text editor
3. Expected: Result text is copied

## Common Issues & Solutions

### Issue 1: Server Won't Start
**Error**: `EADDRINUSE: address already in use :::3000`
**Solution**: 
```bash
# Kill process using port 3000
lsof -ti:3000 | xargs kill -9
# Or use different port
PORT=3001 npm start
```

### Issue 2: SSL Certificate Warning
**Error**: "Your connection is not private"
**Solution**: 
1. Click "Advanced"
2. Click "Proceed to localhost (unsafe)"
3. This is normal for development

### Issue 3: Add-in Won't Load in Office
**Error**: Add-in doesn't appear in ribbon
**Solutions**:
1. Check manifest.xml is accessible: `https://localhost:3000/manifest.xml`
2. Make sure server is running
3. Try refreshing Office app
4. Check Office version (needs Office 365 or Office 2019+)

### Issue 4: API Connection Fails
**Error**: "Failed to authenticate"
**Solutions**:
1. Check internet connection
2. Verify API key is correct
3. Test API directly: `curl https://api.dev2.tens-ai.com/health`

### Issue 5: Text Selection Doesn't Work
**Error**: "No text selected"
**Solutions**:
1. Make sure you have text selected in the Office document
2. Try selecting different text
3. Check that Office app is active

## Testing Checklist

### ✅ Server Tests
- [ ] Server starts without errors
- [ ] Health endpoint responds
- [ ] Web interface loads in browser
- [ ] No JavaScript errors in console

### ✅ Office Integration Tests
- [ ] Add-in installs in Word
- [ ] Add-in installs in Excel
- [ ] Add-in installs in PowerPoint
- [ ] Add-in installs in Outlook
- [ ] Task pane opens correctly
- [ ] Interface displays properly

### ✅ Authentication Tests
- [ ] Invalid API key shows error
- [ ] Valid API key shows success
- [ ] Main interface appears after auth
- [ ] API key is stored securely

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

### ✅ Error Handling Tests
- [ ] Network errors handled gracefully
- [ ] Invalid inputs show appropriate errors
- [ ] Empty content shows validation error
- [ ] API errors display user-friendly messages

## Performance Testing

### Load Test
1. Process multiple requests quickly
2. Check server response times
3. Verify no memory leaks
4. Test with large text inputs

### Stress Test
1. Keep add-in open for extended period
2. Process many AI requests
3. Switch between Office applications
4. Verify stability

## Security Testing

### API Key Security
1. Check API key is not logged in server logs
2. Verify API key is stored securely in browser
3. Test that API key is not exposed in network requests

### Content Security
1. Test with malicious input
2. Verify content is sanitized
3. Check that no sensitive data leaks

## Final Verification

After completing all tests:
1. [ ] All core features work
2. [ ] No critical errors
3. [ ] Performance is acceptable
4. [ ] Security requirements met
5. [ ] User experience is smooth

## Getting Help

If you encounter issues:
1. Check the browser console (F12) for errors
2. Check the server terminal for error messages
3. Verify all prerequisites are installed
4. Test with a fresh Office document
5. Restart the server and try again

For additional support, refer to the main README.md file or contact the development team.

