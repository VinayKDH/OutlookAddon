const express = require('express');
const axios = require('axios');
const router = express.Router();

// TENS AI API configuration
const TENS_AI_API_BASE = process.env.TENS_AI_API_BASE_URL || 'https://api.dev2.tens-ai.com';
const API_TIMEOUT = process.env.TENS_AI_API_TIMEOUT || 30000; // 30 seconds

// Middleware for API key validation
const validateApiKey = (req, res, next) => {
  const apiKey = req.headers['x-api-key'] || req.body.apiKey;
  if (!apiKey) {
    return res.status(401).json({ error: 'API key required' });
  }
  req.apiKey = apiKey;
  next();
};

// Proxy endpoint for TENS AI API calls
router.post('/aimlgyan/*', validateApiKey, async (req, res) => {
  try {
    const endpoint = req.path.replace('/aimlgyan', '');
    const url = `${TENS_AI_API_BASE}${endpoint}`;
    
    const response = await axios({
      method: req.method,
      url: url,
      headers: {
        'Authorization': `Bearer ${req.apiKey}`,
        'Content-Type': 'application/json',
        'User-Agent': 'AI-ML-Gyan-Office-Addon/1.0.0'
      },
      data: req.body,
      timeout: API_TIMEOUT
    });

    res.json(response.data);
  } catch (error) {
    console.error('TENS AI API Error:', error.message);
    
    if (error.response) {
      res.status(error.response.status).json({
        error: 'TENS AI API Error',
        message: error.response.data?.message || error.message
      });
    } else if (error.code === 'ECONNABORTED') {
      res.status(408).json({
        error: 'Request Timeout',
        message: 'The request to TENS AI took too long to complete'
      });
    } else {
      res.status(500).json({
        error: 'Internal Server Error',
        message: 'Failed to connect to TENS AI services'
      });
    }
  }
});

// Text analysis endpoint
router.post('/analyze-text', validateApiKey, async (req, res) => {
  try {
    const { text, analysisType = 'general' } = req.body;
    
    if (!text) {
      return res.status(400).json({ error: 'Text content is required' });
    }

    const response = await axios.post(`${TENS_AI_API_BASE}/analyze`, {
      text: text,
      type: analysisType,
      options: {
        includeSentiment: true,
        includeKeywords: true,
        includeSummary: true
      }
    }, {
      headers: {
        'Authorization': `Bearer ${req.apiKey}`,
        'Content-Type': 'application/json'
      },
      timeout: API_TIMEOUT
    });

    res.json(response.data);
  } catch (error) {
    console.error('Text Analysis Error:', error.message);
    res.status(500).json({
      error: 'Text Analysis Failed',
      message: error.response?.data?.message || error.message
    });
  }
});

// Content generation endpoint
router.post('/generate-content', validateApiKey, async (req, res) => {
  try {
    const { prompt, contentType = 'text', context } = req.body;
    
    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    const response = await axios.post(`${TENS_AI_API_BASE}/generate`, {
      prompt: prompt,
      type: contentType,
      context: context,
      options: {
        maxLength: 1000,
        temperature: 0.7,
        includeMetadata: true
      }
    }, {
      headers: {
        'Authorization': `Bearer ${req.apiKey}`,
        'Content-Type': 'application/json'
      },
      timeout: API_TIMEOUT
    });

    res.json(response.data);
  } catch (error) {
    console.error('Content Generation Error:', error.message);
    res.status(500).json({
      error: 'Content Generation Failed',
      message: error.response?.data?.message || error.message
    });
  }
});

// Document processing endpoint
router.post('/process-document', validateApiKey, async (req, res) => {
  try {
    const { content, documentType, operation } = req.body;
    
    if (!content || !operation) {
      return res.status(400).json({ error: 'Content and operation are required' });
    }

    const response = await axios.post(`${TENS_AI_API_BASE}/process-document`, {
      content: content,
      documentType: documentType,
      operation: operation,
      options: {
        preserveFormatting: true,
        includeSuggestions: true
      }
    }, {
      headers: {
        'Authorization': `Bearer ${req.apiKey}`,
        'Content-Type': 'application/json'
      },
      timeout: API_TIMEOUT
    });

    res.json(response.data);
  } catch (error) {
    console.error('Document Processing Error:', error.message);
    res.status(500).json({
      error: 'Document Processing Failed',
      message: error.response?.data?.message || error.message
    });
  }
});

// Health check for TENS AI services
router.get('/aimlgyan/health', async (req, res) => {
  try {
    const response = await axios.get(`${TENS_AI_API_BASE}/health`, {
      timeout: 5000
    });
    res.json({
      status: 'connected',
      tens_ai: response.data
    });
  } catch (error) {
    res.status(503).json({
      status: 'disconnected',
      error: 'TENS AI services unavailable'
    });
  }
});

module.exports = router;
