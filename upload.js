// upload.js
const express = require('express');
const multer = require('multer');
const bodyParser = require('body-parser');
const spauth = require('node-sp-auth');
const request = require('request-promise-native');
require('dotenv').config();

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
app.use(bodyParser.json());

// Environment variables for credentials and config
const SP_SITE_URL = process.env.SP_SITE_URL || 'https://www.ifwg.co.za';
const SP_LIST_NAME = process.env.SP_LIST_NAME || 'Enquiry Details';
const SP_LIBRARY_NAME = process.env.SP_LIBRARY_NAME || 'EnquiryFormDocuments';
const SP_USERNAME = process.env.SP_USERNAME || 'svc_IFWGEnquiry';
const SP_PASSWORD = process.env.SP_PASSWORD || 'sQDrLej^[7yVSR\'TxZ';
const SP_DOMAIN = process.env.SP_DOMAIN || 'FSCA';

// Helper: Get SharePoint Auth Headers
async function getAuthHeaders(siteUrl) {
  return spauth.getAuth(siteUrl, {
    username: SP_USERNAME,
    password: SP_PASSWORD,
    domain: SP_DOMAIN
  });
}

// Helper: Create List Item
app.post('/api/create-list-item', async (req, res) => {
  try {
    const item = req.body;
    const auth = await getAuthHeaders(SP_SITE_URL);
    const endpoint = `${SP_SITE_URL}/_api/web/lists/getbytitle('${SP_LIST_NAME}')/items`;
    const payload = {
      __metadata: { type: 'SP.Data.Enquiry_x0020_DetailsListItem' },
      ...item
    };
    const options = {
      url: endpoint,
      method: 'POST',
      headers: {
        ...auth.headers,
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      },
      body: JSON.stringify(payload),
      json: false
    };
    const response = await request(options);
    const responseData = JSON.parse(response);
    res.status(201).json(responseData.d);
  } catch (error) {
    console.error('Error creating list item:', error.message);
    res.status(500).json({ error: error.message, details: error.error });
  }
});

// Helper: Upload File
app.post('/api/upload-file', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    const folderName = req.body.folderName || '';
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    const auth = await getAuthHeaders(SP_SITE_URL);
    const fileName = folderName ? `${folderName}_${file.originalname}` : file.originalname;
    const uploadUrl = `${SP_SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${SP_LIBRARY_NAME}')/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`;
    const options = {
      url: uploadUrl,
      method: 'POST',
      headers: {
        ...auth.headers,
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/octet-stream'
      },
      body: file.buffer,
      json: false
    };
    const response = await request(options);
    const responseData = JSON.parse(response);
    res.status(200).json(responseData.d);
  } catch (error) {
    console.error('Error uploading file:', error.message);
    res.status(500).json({ error: error.message, details: error.error });
  }
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`SharePoint upload API running on port ${PORT}`);
});