# Troubleshooting Guide - TENS AI Office Add-in

## 🚨 Issue: Localhost Taskpane Doesn't Open

### Quick Fix (HTTP Version - Recommended for Development)

**Step 1: Stop any running servers**
```bash
# Kill any processes on port 3000
lsof -ti:3000 | xargs kill -9
```

**Step 2: Start HTTP server**
```bash
node server-http.js
```

**Step 3: Test in browser**
- Open: `http://localhost:3000/taskpane.html`
- Should see the TENS AI interface

**Step 4: Install in Office**
- Use `manifest-http.xml` instead of `manifest.xml`
- Upload the HTTP manifest to Office

### Alternative Fix (HTTPS Version)

**Step 1: Generate SSL certificates**
```bash
node setup-ssl.js
```

**Step 2: Start HTTPS server**
```bash
npm start
```

**Step 3: Test in browser**
- Open: `https://localhost:3000/taskpane.html`
- Click "Advanced" → "Proceed to localhost" when you see SSL warning

## 🔍 Common Issues & Solutions

### Issue 1: Server Won't Start

**Error**: `EADDRINUSE: address already in use :::3000`

**Solution**:
```bash
# Find and kill process using port 3000
lsof -ti:3000 | xargs kill -9

# Or use a different port
PORT=3001 node server-http.js
```

### Issue 2: SSL Certificate Errors

**Error**: `SSL routines:ST_CONNECT:tlsv1 alert protocol version`

**Solution**:
```bash
# Use HTTP version instead
node server-http.js

# Or generate SSL certificates
node setup-ssl.js
npm start
```

### Issue 3: Browser Shows "Connection Refused"

**Error**: `ERR_CONNECTION_REFUSED`

**Solution**:
1. Check if server is running: `lsof -i :3000`
2. Start server: `node server-http.js`
3. Check server logs for errors

### Issue 4: Office Add-in Won't Load

**Error**: Add-in doesn't appear in ribbon

**Solutions**:
1. **Check manifest accessibility**:
   - HTTP: `http://localhost:3000/manifest-http.xml`
   - HTTPS: `https://localhost:3000/manifest.xml`

2. **Verify server is running**:
   ```bash
   curl http://localhost:3000/health
   ```

3. **Check Office version**:
   - Requires Office 365 or Office 2019+
   - Update Office if needed

4. **Try different manifest**:
   - Use `manifest-http.xml` for HTTP server
   - Use `manifest.xml` for HTTPS server

### Issue 5: Taskpane Shows Blank Page

**Error**: White/blank page in taskpane

**Solutions**:
1. **Check browser console** (F12):
   - Look for JavaScript errors
   - Check network requests

2. **Verify static files**:
   ```bash
   curl http://localhost:3000/taskpane.html
   curl http://localhost:3000/taskpane.js
   curl http://localhost:3000/styles.css
   ```

3. **Check file permissions**:
   ```bash
   ls -la public/
   ```

### Issue 6: API Connection Fails

**Error**: "Failed to authenticate" or API errors

**Solutions**:
1. **Test API connectivity**:
   ```bash
   curl http://localhost:3000/api/aimlgyan/health
   ```

2. **Check dev2.tens-ai.com connectivity**:
   ```bash
   curl -I https://www.dev2.tens-ai.com
   curl -I https://api.dev2.tens-ai.com
   ```

3. **Verify API key**:
   - Get key from: `https://www.dev2.tens-ai.com/api-keys`
   - Test with valid key

### Issue 7: CORS Errors

**Error**: `Access to fetch at '...' from origin '...' has been blocked by CORS policy`

**Solution**:
1. Check server CORS configuration
2. Ensure all domains are whitelisted
3. Restart server after changes

## 🛠️ Debug Commands

### Check Server Status
```bash
# Check if server is running
lsof -i :3000

# Test server health
curl http://localhost:3000/health

# Check server logs
# Look at terminal where server is running
```

### Check File Accessibility
```bash
# Test main files
curl http://localhost:3000/taskpane.html
curl http://localhost:3000/manifest-http.xml
curl http://localhost:3000/assets/logo.svg

# Check file permissions
ls -la public/
ls -la public/assets/
```

### Check Network Connectivity
```bash
# Test localhost
ping localhost

# Test external APIs
curl -I https://www.dev2.tens-ai.com
curl -I https://api.dev2.tens-ai.com

# Test with timeout
curl --connect-timeout 5 http://localhost:3000/health
```

## 🔧 Development Workflow

### For HTTP Development (Recommended)
```bash
# 1. Start HTTP server
node server-http.js

# 2. Test in browser
open http://localhost:3000/taskpane.html

# 3. Install in Office
# Use manifest-http.xml

# 4. Test add-in functionality
```

### For HTTPS Development
```bash
# 1. Generate SSL certificates (one-time setup)
node setup-ssl.js

# 2. Start HTTPS server
npm start

# 3. Test in browser
open https://localhost:3000/taskpane.html

# 4. Install in Office
# Use manifest.xml

# 5. Test add-in functionality
```

## 📋 Testing Checklist

### ✅ Server Tests
- [ ] Server starts without errors
- [ ] Health endpoint responds: `curl http://localhost:3000/health`
- [ ] Taskpane loads in browser: `http://localhost:3000/taskpane.html`
- [ ] Manifest is accessible: `http://localhost:3000/manifest-http.xml`
- [ ] Static files load correctly

### ✅ Office Integration Tests
- [ ] Add-in installs using correct manifest
- [ ] Taskpane opens in Office
- [ ] Interface displays properly
- [ ] No JavaScript errors in browser console

### ✅ API Tests
- [ ] API health check works
- [ ] External API connectivity works
- [ ] Authentication works with valid API key
- [ ] Error handling works for invalid keys

## 🆘 Getting Help

### If Nothing Works
1. **Restart everything**:
   ```bash
   # Kill all processes
   lsof -ti:3000 | xargs kill -9
   
   # Clear npm cache
   npm cache clean --force
   
   # Reinstall dependencies
   rm -rf node_modules package-lock.json
   npm install
   
   # Start fresh
   node server-http.js
   ```

2. **Check system requirements**:
   - Node.js version 16+
   - Office 365 or Office 2019+
   - Modern browser (Chrome, Firefox, Safari, Edge)

3. **Try different approaches**:
   - Use HTTP instead of HTTPS
   - Try different port (3001, 3002, etc.)
   - Test in different Office applications
   - Test in Office Online vs Desktop

### Contact Support
- Check server logs for specific error messages
- Test with minimal setup (HTTP server only)
- Verify all prerequisites are installed
- Document exact error messages and steps to reproduce

## 🎯 Success Indicators

The add-in is working correctly when:
- ✅ Server starts without errors
- ✅ Browser shows taskpane interface
- ✅ Office loads the add-in successfully
- ✅ Taskpane opens in Office applications
- ✅ API connections work (with valid keys)
- ✅ All features function as expected

Remember: For development, the HTTP version (`server-http.js` + `manifest-http.xml`) is often easier to work with than HTTPS!
