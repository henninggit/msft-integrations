const https = require('https');
const fs = require('fs');
const path = require('path');

const certPath = path.join(require('os').homedir(), '.office-addin-dev-certs', 'localhost.crt');
const keyPath = path.join(require('os').homedir(), '.office-addin-dev-certs', 'localhost.key');
const caPath = path.join(require('os').homedir(), '.office-addin-dev-certs', 'ca.crt');

console.log('Certificate paths:');
console.log('Cert:', certPath, 'exists:', fs.existsSync(certPath));
console.log('Key:', keyPath, 'exists:', fs.existsSync(keyPath));
console.log('CA:', caPath, 'exists:', fs.existsSync(caPath));

if (fs.existsSync(certPath)) {
  const cert = fs.readFileSync(certPath, 'utf8');
  console.log('\nCertificate content preview:');
  console.log(cert.substring(0, 200) + '...');
}

// Test HTTPS request to ourselves
const options = {
  hostname: 'localhost',
  port: 3000,
  path: '/manifest.xml',
  method: 'GET',
  rejectUnauthorized: false // Allow self-signed for testing
};

console.log('\nTesting HTTPS connection...');
const req = https.request(options, (res) => {
  console.log('Status Code:', res.statusCode);
  console.log('Headers:', res.headers);
  
  let data = '';
  res.on('data', (chunk) => {
    data += chunk;
  });
  
  res.on('end', () => {
    console.log('Response length:', data.length);
    console.log('First 200 chars:', data.substring(0, 200));
  });
});

req.on('error', (e) => {
  console.error('Request error:', e.message);
  console.error('Error code:', e.code);
});

req.end();
