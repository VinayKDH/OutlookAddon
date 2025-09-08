#!/usr/bin/env node

const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

console.log('🔐 Setting up SSL certificates for development...\n');

const sslDir = path.join(__dirname, 'ssl');
const certFile = path.join(sslDir, 'cert.pem');
const keyFile = path.join(sslDir, 'key.pem');

// Create SSL directory if it doesn't exist
if (!fs.existsSync(sslDir)) {
  fs.mkdirSync(sslDir);
  console.log('✅ Created SSL directory');
}

// Check if certificates already exist
if (fs.existsSync(certFile) && fs.existsSync(keyFile)) {
  console.log('✅ SSL certificates already exist');
  console.log('📁 Certificate:', certFile);
  console.log('📁 Private key:', keyFile);
  console.log('\n🚀 You can now run: npm start (for HTTPS)');
  process.exit(0);
}

// Generate self-signed certificate
try {
  console.log('🔧 Generating self-signed SSL certificate...');
  
  const command = `openssl req -x509 -newkey rsa:4096 -keyout "${keyFile}" -out "${certFile}" -days 365 -nodes -subj "/C=US/ST=State/L=City/O=Organization/CN=localhost"`;
  
  execSync(command, { stdio: 'inherit' });
  
  console.log('✅ SSL certificate generated successfully!');
  console.log('📁 Certificate:', certFile);
  console.log('📁 Private key:', keyFile);
  console.log('\n🚀 You can now run: npm start (for HTTPS)');
  console.log('⚠️  Note: Browsers will show a security warning for self-signed certificates');
  console.log('   Click "Advanced" → "Proceed to localhost" to continue');
  
} catch (error) {
  console.error('❌ Failed to generate SSL certificate:', error.message);
  console.log('\n💡 Alternative: Use the HTTP version for development:');
  console.log('   node server-http.js');
  console.log('   Then use manifest-http.xml in Office');
  process.exit(1);
}
