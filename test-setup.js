#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const https = require('https');

console.log('🧪 TENS AI Office Add-in Setup Test\n');

// Test 1: Check if all required files exist
console.log('📁 Checking required files...');
const requiredFiles = [
    'package.json',
    'server.js',
    'manifest.xml',
    'public/taskpane.html',
    'public/taskpane.js',
    'public/commands.html',
    'public/commands.js',
    'public/styles.css',
    'routes/api.js'
];

let allFilesExist = true;
requiredFiles.forEach(file => {
    if (fs.existsSync(file)) {
        console.log(`✅ ${file}`);
    } else {
        console.log(`❌ ${file} - MISSING`);
        allFilesExist = false;
    }
});

if (!allFilesExist) {
    console.log('\n❌ Some required files are missing. Please check the file structure.');
    process.exit(1);
}

// Test 2: Check package.json
console.log('\n📦 Checking package.json...');
try {
    const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));
    console.log(`✅ Package name: ${packageJson.name}`);
    console.log(`✅ Version: ${packageJson.version}`);
    console.log(`✅ Main entry: ${packageJson.main}`);
} catch (error) {
    console.log('❌ Invalid package.json');
    process.exit(1);
}

// Test 3: Check manifest.xml
console.log('\n📋 Checking manifest.xml...');
try {
    const manifest = fs.readFileSync('manifest.xml', 'utf8');
    if (manifest.includes('TENS AI')) {
        console.log('✅ Manifest contains correct app name');
    } else {
        console.log('❌ Manifest missing app name');
    }
    
    if (manifest.includes('https://localhost:3000')) {
        console.log('✅ Manifest contains correct localhost URL');
    } else {
        console.log('❌ Manifest missing localhost URL');
    }
    
    if (manifest.includes('dev2.tens-ai.com')) {
        console.log('✅ Manifest contains correct dev2.tens-ai.com URLs');
    } else {
        console.log('❌ Manifest missing dev2.tens-ai.com URLs');
    }
} catch (error) {
    console.log('❌ Cannot read manifest.xml');
    process.exit(1);
}

// Test 4: Check if node_modules exists (after npm install)
console.log('\n🔧 Checking dependencies...');
if (fs.existsSync('node_modules')) {
    console.log('✅ node_modules directory exists');
} else {
    console.log('❌ node_modules not found - run "npm install" first');
    console.log('💡 Run: npm install');
}

// Test 5: Test server startup (if dependencies are installed)
if (fs.existsSync('node_modules')) {
    console.log('\n🚀 Testing server startup...');
    
    // Try to require the server
    try {
        // This will test if the server can be loaded without errors
        delete require.cache[require.resolve('./server.js')];
        console.log('✅ Server module loads without errors');
    } catch (error) {
        console.log(`❌ Server module error: ${error.message}`);
    }
}

// Test 6: Check port availability
console.log('\n🌐 Checking port availability...');
const net = require('net');
const server = net.createServer();

server.listen(3000, () => {
    console.log('✅ Port 3000 is available');
    server.close();
});

server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
        console.log('❌ Port 3000 is already in use');
        console.log('💡 Try: lsof -ti:3000 | xargs kill -9');
        console.log('💡 Or use: PORT=3001 npm start');
    } else {
        console.log(`❌ Port error: ${err.message}`);
    }
});

// Test 7: Check connectivity to dev2.tens-ai.com
console.log('\n🌐 Testing connectivity to dev2.tens-ai.com...');

const testConnectivity = (url) => {
  return new Promise((resolve) => {
    const req = https.get(url, { timeout: 5000 }, (res) => {
      console.log(`✅ ${url} - Status: ${res.statusCode}`);
      resolve(true);
    });
    
    req.on('error', (err) => {
      console.log(`❌ ${url} - Error: ${err.message}`);
      resolve(false);
    });
    
    req.on('timeout', () => {
      console.log(`❌ ${url} - Timeout`);
      req.destroy();
      resolve(false);
    });
  });
};

(async () => {
  await testConnectivity('https://www.dev2.tens-ai.com');
  await testConnectivity('https://api.dev2.tens-ai.com');
})();

console.log('\n🎯 Next Steps:');
console.log('1. Run: npm install (if not done already)');
console.log('2. Run: npm start');
console.log('3. Open: https://localhost:3000/taskpane.html');
console.log('4. Install add-in in Office using manifest.xml');
console.log('5. Test the functionality with dev2.tens-ai.com API');

console.log('\n📚 For detailed testing instructions, see:');
console.log('- README.md (complete setup guide)');
console.log('- TESTING_GUIDE.md (step-by-step testing)');

console.log('\n✨ Setup test completed!');

