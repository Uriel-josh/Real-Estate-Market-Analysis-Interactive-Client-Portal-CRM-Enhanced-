# Real-Estate-Market-Analysis-Interactive-Client-Portal-CRM-Enhanced
// Microsoft 365 Copilot Integration for CRM Analysis
// This code integrates with Microsoft 365 APIs to analyze client interactions

const axios = require('axios');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
require('isomorphic-fetch');
const { TextAnalyticsClient, AzureKeyCredential } = require('@azure/ai-text-analytics');

// Configuration settings for Microsoft 365 integration
const config = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel, message, containsPii) {
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: msal.LogLevel.Verbose,
    }
  }
};

// Initialize MSAL application for authentication
const cca = new msal.ConfidentialClientApplication(config);

// Text Analytics Configuration
const textAnalyticsClient = new TextAnalyticsClient(
  process.env.AZURE_TEXT_ANALYTICS_ENDPOINT,
  new AzureKeyCredential(process.env.AZURE_TEXT_ANALYTICS_KEY)
);

// Function to get access token for Microsoft Graph API
async function getToken() {
  const tokenRequest = {
    scopes: ['https://graph.microsoft.com/.default'],
  };

  try {
    const response = await cca.acquireTokenByClientCredential(tokenRequest);
    return response.accessToken;
  } catch (error) {
    console.error('Error acquiring token:', error);
    throw error;
  }
}

// Initialize Microsoft Graph Client
async function getGraphClient() {
  const accessToken = await getToken();
  
  // Initialize Graph client with authentication provider
  const authProvider = (callback) => {
    callback(null, accessToken);
  };
  
  const graphClient = Client.init({
    authProvider: authProvider,
  });
  
  return graphClient;
}

// Function to fetch and analyze client emails
async function analyzeClientEmails(clientEmail, timeRange = 30) {
  try {
    const graphClient = await getGraphClient();
    
    // Calculate date range for filtering
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - timeRange);
    
    const startDateString = startDate.toISOString();
    const endDateString = endDate.toISOString();
    
    // Search for messages related to the client
    const messagesResponse = await graphClient
      .api('/me/messages')
      .filter(`receivedDateTime ge ${startDateString} and receivedDateTime le ${endDateString} and (from/emailAddress/address eq '${clientEmail}' or toRecipients/any(r:r/emailAddress/address eq '${clientEmail}'))`)
      .select('subject,bodyPreview,receivedDateTime,from,importance')
      .orderBy('receivedDateTime desc')
      .top(100)
      .get();
    
    // Extract email content for analysis
    const emailContents = messagesResponse.value.map(message => ({
      id: message.id,
      subject: message.subject,
      content: message.bodyPreview,
      date: message.receivedDateTime,
      from: message.from.emailAddress.address,
      importance: message.importance
    }));
    
    // Perform sentiment analysis
    const sentiments = await analyzeSentiment(emailContents.map(email => email.content));
    
    // Combine email data with sentiment analysis
    const analyzedEmails = emailContents.map((email, index) => ({
      ...email,
      sentiment: sentiments[index].sentiment,
      confidenceScores: sentiments[index].confidenceScores
    }));
    
    // Perform key phrase extraction on all emails combined
    const allContent = emailContents.map(email => email.content).join(' ');
    const keyPhrases = await extractKeyPhrases([allContent]);
    
    // Get overall sentiment trend
    const sentimentTrend = calculateSentimentTrend(analyzedEmails);
    
    return {
      emails: analyzedEmails,
      keyPhrases: keyPhrases[0],
      sentimentTrend,
      interactionCount: emailContents.length
    };
  } catch (error) {
    console.error('Error analyzing client emails:', error);
    throw error;
  }
}

// Function to perform sentiment analysis on text
async function analyzeSentiment(texts) {
  const results = [];
  
  try {
    const sentimentResults = await textAnalyticsClient.analyzeSentiment(texts);
    
    sentimentResults.forEach(result => {
      if (result.error) {
        results.push({
          sentiment: 'unknown',
          confidenceScores: { positive: 0, neutral: 0, negative: 0 }
        });
      } else {
        results.push({
          sentiment: result.sentiment,
          confidenceScores: result.confidenceScores
        });
      }
    });
  } catch (error) {
    console.error('Error in sentiment analysis:', error);
  }
  
  return results;
}

// Function to extract key phrases from text
async function extractKeyPhrases(texts) {
  const results = [];
  
  try {
    const keyPhraseResults = await textAnalyticsClient.extractKeyPhrases(texts);
    
    keyPhraseResults.forEach(result => {
      if (result.error) {
        results.push([]);
      } else {
        results.push(result.keyPhrases);
      }
    });
  } catch (error) {
    console.error('Error in key phrase extraction:', error);
  }
  
  return results;
}

// Function to calculate sentiment trend over time
function calculateSentimentTrend(analyzedEmails) {
  if (analyzedEmails.length === 0) {
    return { trend: 'neutral', score: 0 };
  }
  
  // Sort emails by date
  const sortedEmails = [...analyzedEmails].sort((a, b) => 
    new Date(a.date) - new Date(b.date)
  );
  
  // Convert sentiment to numerical score for each email
  const sentimentScores = sortedEmails.map(email => {
    const { positive, neutral, negative } = email.confidenceScores;
    return positive - negative; // Simple score calculation
  });
  
  // Calculate overall trend using linear regression
  const n = sentimentScores.length;
  
  if (n < 2) {
    return {
      trend: sentimentScores[0] > 0 ? 'positive' : sentimentScores[0] < 0 ? 'negative' : 'neutral',
      score: sentimentScores[0] || 0
    };
  }
  
  const indices = Array.from({ length: n }, (_, i) => i);
  
  // Calculate slope using least squares method
  const sumX = indices.reduce((sum, x) => sum + x, 0);
  const sumY = sentimentScores.reduce((sum, y) => sum + y, 0);
  const sumXY = indices.reduce((sum, x, i) => sum + x * sentimentScores[i], 0);
  const sumXX = indices.reduce((sum, x) => sum + x * x, 0);
  
  const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
  const averageScore = sumY / n;
  
  // Determine trend based on slope
  let trend;
  if (slope > 0.1) {
    trend = 'improving';
  } else if (slope < -0.1) {
    trend = 'declining';
  } else {
    trend = averageScore > 0.2 ? 'consistently positive' : 
            averageScore < -0.2 ? 'consistently negative' : 'neutral';
  }
  
  return { trend, score: averageScore, slope };
}

// Function to extract client preferences from email content
async function extractClientPreferences(clientId) {
  try {
    // Get client email from database
    const { Pool } = require('pg');
    const pool = new Pool();
    
    const clientQuery = await pool.query(
      'SELECT email FROM clients WHERE client_id = $1',
      [clientId]
    );
    
    if (clientQuery.rows.length === 0) {
      throw new Error('Client not found');
    }
    
    const clientEmail = clientQuery.rows[0].email;
    
    // Analyze client emails
    const emailAnalysis = await analyzeClientEmails(clientEmail, 90); // Last 90 days
    
    // Extract key phrases related to real estate
    const realEstateKeywords = [
      'bedroom', 'bathroom', 'square footage', 'backyard', 'garage',
      'kitchen', 'neighborhood', 'school', 'price', 'budget',
      'location', 'commute', 'single family', 'condo', 'townhouse',
      'apartment', 'view', 'basement', 'yard', 'pool', 'renovation',
      'new construction', 'open floor plan', 'modern', 'traditional'
    ];
    
    // Filter key phrases related to real estate
    const relevantPhrases = emailAnalysis.keyPhrases.filter(phrase => 
      realEstateKeywords.some(keyword => 
        phrase.toLowerCase().includes(keyword.toLowerCase())
      )
    );
    
    // Call Microsoft Copilot API to analyze preferences
    const copilotAnalysis = await analyzeCopilot(
      clientEmail, 
      relevantPhrases,
      emailAnalysis.emails.map(e => e.content).join('\n\n')
    );
    
    // Update client preferences in database
    if (copilotAnalysis.preferences) {
      await updateClientPreferences(clientId, copilotAnalysis.preferences);
    }
    
    return {
      clientId,
      relevantPhrases,
      preferences: copilotAnalysis.preferences,
      sentimentTrend: emailAnalysis.sentimentTrend
    };
  } catch (error) {
    console.error('Error extracting client preferences:', error);
    throw error;
  }
}

// Function to call Microsoft Copilot to extract and analyze preferences
async function analyzeCopilot(clientEmail, keyPhrases, emailContent) {
  try {
    // Setup Copilot API request
    const token = await getToken();
    
    const promptTemplate = `
    Analyze the following key phrases and email content from real estate client communications.
    Extract specific client preferences related to:
    1. Property type preferences (e.g., single family, condo, townhouse)
    2. Size preferences (bedrooms, bathrooms, square footage)
    3. Location preferences (neighborhoods, school districts)
    4. Price range
    5. Must-have features
    6. Nice-to-have features
    7. Deal-breaker features
    
    Key phrases: ${keyPhrases.join(', ')}
    
    Sample email content:
    ${emailContent.substring(0, 2000)}
    
    Format your response as a structured JSON object with these preference categories.
    `;
    
    const response = await axios({
      method: 'POST',
      url: `https://graph.microsoft.com/v1.0/me/onlineMeetings/cognitiveServices/copilot/analyzeContent`,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      data: {
        prompt: promptTemplate,
        temperature: 0.1,
        maxTokens: 800
      }
    });
    
    // Parse response
    let preferences = {};
    
    try {
      const responseText = response.
