const express = require('express');
const cors = require('cors');
const generateDocumentRoute = require('../routes/generateDocument');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());

// Routes
app.use('/generate-document', generateDocumentRoute);

// Basic health check route
app.get('/', (req, res) => {
  res.json({ status: 'Document Generation Service is running' });
});

// For Vercel serverless deployment
module.exports = app;

// For local development
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
  });
}
