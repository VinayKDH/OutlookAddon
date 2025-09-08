const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const compression = require('compression');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Security middleware
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com"],
      fontSrc: ["'self'", "https://fonts.gstatic.com"],
      scriptSrc: ["'self'", "https://appsforoffice.microsoft.com"],
      connectSrc: ["'self'", "https://www.dev2.tens-ai.com", "https://api.dev2.tens-ai.com"],
      imgSrc: ["'self'", "data:", "https:"],
      frameSrc: ["'self'", "https://appsforoffice.microsoft.com"]
    }
  }
}));

// CORS configuration
app.use(cors({
  origin: [
    'http://localhost:3000',
    'http://127.0.0.1:3000',
    'https://localhost:3000',
    'https://127.0.0.1:3000',
    'https://www.dev2.tens-ai.com',
    'https://api.dev2.tens-ai.com'
  ],
  credentials: true
}));

// Compression and parsing middleware
app.use(compression());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Serve static files
app.use(express.static(path.join(__dirname, 'public')));

// API Routes
app.use('/api', require('./routes/api'));

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ 
    status: 'healthy', 
    timestamp: new Date().toISOString(),
    version: '1.0.0'
  });
});

// Serve the main taskpane
app.get('/taskpane.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'taskpane.html'));
});

// Serve the commands page
app.get('/commands.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'commands.html'));
});

// Serve the manifest
app.get('/manifest.xml', (req, res) => {
  res.setHeader('Content-Type', 'application/xml');
  res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Error:', err);
  res.status(500).json({ 
    error: 'Internal server error',
    message: process.env.NODE_ENV === 'development' ? err.message : 'Something went wrong'
  });
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({ error: 'Not found' });
});

// Start server
app.listen(PORT, () => {
  console.log(`TENS AI Office Add-in server running on port ${PORT}`);
  console.log(`Taskpane: http://localhost:${PORT}/taskpane.html`);
  console.log(`Manifest: http://localhost:${PORT}/manifest.xml`);
  console.log(`Health: http://localhost:${PORT}/health`);
});

module.exports = app;
