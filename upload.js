const express = require('express');
const fileUpload = require('express-fileupload');
const { execFile } = require('child_process');
const path = require('path');
const fs = require('fs');
const cors = require('cors')
 
require('dotenv').config();
 
const app = express();
 
app.use(cors({
    origin: 'https://www.ifwg.co.za',
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: [
        'Origin',
        'X-Requested-With',
        'Content-Type',
        'Accept',
        'Authorization',
        'X-RequestDigest',
        'Cache-Control',
        'Pragma'
        ]
  }));
 
  app.use((req, res, next) => {
    res.header('Vary', 'Origin');
    next();
  });
 
 
app.options('*', cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
 
app.set('trust proxy', true); // Trust X-Forwarded-* headers
app.set('x-powered-by', false); // Remove Express header
app.use(fileUpload({
useTempFiles: true,
tempFileDir: '/tmp/',
createParentPath: true
}));
 
// Add request logging
app.use((req, res, next) => {
console.log(`${new Date().toISOString()} - ${req.method} ${req.url}`);
console.log('Headers:', req.headers);
next();
});
const siteUrl = process.env.SP_SITE_URL;
const domain = process.env.SP_DOMAIN;
const username = process.env.SP_USERNAME;
const password = process.env.SP_PASSWORD;
const listName = process.env.SP_LIST_NAME;
const libraryName = process.env.SP_LIBRARY_NAME;
 
function formatDateToString(date) {
// Ensure the input is a valid Date object
if (!(date instanceof Date) || isNaN(date)) {
console.error("Invalid input: Please provide a valid Date object.");
return ""; // Return an empty string or throw an error based on desired behavior
}
const d = date.getDate(); // Day of the month (1-31)
const M = date.getMonth() + 1; // Month (0-11, so add 1 for 1-12)
const yyyy = date.getFullYear(); // Full year (e.g., 2023)
const H = date.getHours(); // Hour (0-23)
const m = date.getMinutes(); // Minute (0-59)
const s = date.getSeconds(); // Second (0-59)
const SSS = date.getMilliseconds(); // Millisecond (0-999)
// Construct the string. No leading zeros are added for single-digit numbers
// as per the 'dMyyyyHmsSSS' format.
return `${d}${M}${yyyy}${H}${m}${s}${SSS}`;
}
// Health endpoints
 
app.get('/api/health', (req, res) => {
res.json({
status: 'ok',
timestamp: new Date().toISOString(),
server: 'Node.js API'
});
});
 
// Add List Item
app.post('/api/create-list-item', (req, res) => {
const fields = [
'Title','FullName','OrganisationName','ContactNumber','EmailAddress','WebsiteAddress','OperationLocation',
'CountriesOfOperation','OperationLength','PrimaryBusinessAreas','ProductServiceCategory','OtherProductServiceCategory',
'OperationalStatus','RegulatoryStatus','Regulators','OtherRegulator','ProductServiceDescription','Questions',
'AdditionalInformation','FAQConfirmation','ConsentConfirmation','FileAttachments','NumberOfAttachments','SubmissionDate'
];
const args = [
'-File', 'add-list-item.ps1',
'-siteUrl', siteUrl,
'-username', username,
'-password', password,
'-domain', domain,
'-listName', listName
];
fields.forEach(f => {
args.push('-' + f, req.body[f] || '');
});
execFile('powershell.exe', args, (error, stdout, stderr) => {
console.log('PowerShell stdout:', stdout);
console.log('PowerShell stderr:', stderr);
console.log('PowerShell error:', error);
if (error) {
console.error('PowerShell execution failed:', error);
return res.status(500).json({ error: stderr || error.message });
}
try {
const data = JSON.parse(stdout.trim());
console.log('Parsed PowerShell output:', data);
if (data && data.Id && data.Id > 0) {
return res.json(data);
} else {
console.error('PowerShell output missing valid Id:', data);
return res.status(500).json({ error: 'PowerShell did not return a valid item ID', details: stdout.trim() });
}
} catch (e) {
console.error('JSON parsing failed:', e);
console.error('Raw stdout:', stdout);
return res.status(500).json({ error: 'PowerShell output was not valid JSON', details: stdout.trim() });
}
});
});
// Upload Document
app.post('/api/upload-doc', async (req, res) => {
if (!req.files?.document) {
return res.status(400).json({ error: 'No file uploaded' });
}
console.log('=== File Upload Debug ===');
console.log('Headers received:', req.headers);
console.log('File received:', req.files.document ? {
originalname: req.files.document.name,
size: req.files.document.size,
mimetype: req.files.document.mimetype
} : 'No file');
console.log('Body fields:', req.body);
const file = req.files.document;
const datep = formatDateToString(new Date());
file.name = `photo_specimen${datep}${path.parse(file.name).ext}`;
file.mv(`${process.env.FILE_UPLOAD_PATH}/${file.name}`, async (err) => {
if (err) {
return next(new ErrorResponse(`Problem with file upload ${err}`, 500));
}
});
const tempFilePath = `${process.env.FILE_UPLOAD_PATH}/${file.name}`;
console.log('tempFilePath',tempFilePath)// Save file temporarily
//await file.mv(tempFilePath);
console.log('siteUrl',siteUrl);
execFile('powershell.exe', [
'-File', 'sp-upload.ps1',
'-siteUrl', siteUrl,
'-libraryName', libraryName,
'-filePath', tempFilePath,
'-domain', domain,
'-username', username,
'-password', password
], (error, stdout, stderr) => {
fs.unlinkSync(tempFilePath);
if (error) {
console.error('PowerShell upload error:', error);
console.error('PowerShell stderr:', stderr);
return res.status(500).json({ error: stderr || error.message });
}
try {
const data = JSON.parse(stdout.trim());
console.log('Upload PowerShell output:', data);
if (data && data.ServerRelativeUrl) {
return res.json(data);
} else {
console.log(data);
return res.status(500).json({ error: 'PowerShell did not return a valid file URL', details: stdout.trim() });
}
} catch (e) {
console.error('Upload JSON parsing failed:', e);
console.error('Raw upload stdout:', stdout);
return res.status(500).json({ error: 'PowerShell output was not valid JSON', details: stdout.trim() });
}
});
});
app.listen(3001, () => {
console.log('SharePoint upload API running on port 3001');
});