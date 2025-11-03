// Test script to create Office file directly
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');

async function testCreateOfficeFile() {
    // Replace with your actual values
    const tenantId = 'your-tenant-id';
    const clientId = 'your-client-id';
    const clientSecret = 'your-client-secret';
    
    try {
        console.log('🔐 Authenticating...');
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        
        const graphClient = Client.initWithMiddleware({
            authProvider: {
                getAccessToken: async () => {
                    const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
                    return tokenResponse.token;
                }
            }
        });
        
        console.log('📄 Creating Word document...');
        
        // Create simple Word content
        const wordContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is MyDocument created by Lokka MCP test script.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
        
        // Upload to OneDrive
        const uploadResponse = await graphClient
            .api('/me/drive/root:/MyDocument.docx:/content')
            .header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            .put(Buffer.from(wordContent));
            
        console.log('✅ Success!');
        console.log('File ID:', uploadResponse.id);
        console.log('File URL:', uploadResponse.webUrl);
        console.log('Created by:', uploadResponse.createdBy?.user?.displayName || 'SharePoint App');
        
    } catch (error) {
        console.error('❌ Error:', error.message);
    }
}

testCreateOfficeFile();
