# Word Integration Office Add-in

A simple Microsoft Word Office Add-in that demonstrates integration capabilities using Office.js APIs.

## Features

- **Insert Text**: Add sample text to the document
- **Get Selection**: Retrieve currently selected text
- **Format Text**: Apply bold, color, and size formatting to selected text
- **Insert Table**: Create a formatted table with sample data
- **Document Info**: Get statistics about the document (paragraphs, tables, word count)

## Prerequisites

- Microsoft Word (Desktop, Online, or Mac)
- Node.js (version 16 or higher)
- npm or yarn

## Setup Instructions

1. **Install dependencies**:
   ```bash
   npm install
   ```

2. **Start the development environment** (requires 2 terminals):

   **Terminal 1 - Start Webpack Dev Server:**
   ```bash
   npx webpack serve --mode development
   ```
   This starts the dev server at `https://localhost:3000/` and must keep running.

   **Terminal 2 - Launch Word with Add-in:**
   ```bash
   npm start
   ```
   This will:
   - Install SSL certificates for local development
   - Launch Word with the add-in loaded
   - The add-in will appear in the **Home** tab under "Word Integration"

3. **Fix SSL Certificate Issues** (if needed):
   - Open your browser and go to `https://localhost:3000/taskpane.html`
   - Accept the certificate warning ("Advanced" → "Proceed to localhost")
   - Return to Word - the add-in should now work properly

4. **Alternative Manual Setup**:
   - Open Microsoft Word
   - Go to **Insert** > **Add-ins** > **Upload My Add-in**
   - Browse and select the `manifest.xml` file from this project

## Development

- **Development mode**: `npm run dev` - Runs webpack in watch mode
- **Build for production**: `npm run build`
- **Validate manifest**: `npm run validate`

## Project Structure

```
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # Main UI
│   │   └── taskpane.js      # Word integration logic
│   └── commands/
│       └── commands.html    # Command functions
├── assets/
│   └── icon-16.svg         # Add-in icons
├── manifest.xml            # Add-in configuration
├── webpack.config.js       # Build configuration
└── package.json           # Dependencies and scripts
```

## How It Works

This add-in uses the **Office.js JavaScript API** to interact with Word:

1. **Office.onReady()**: Initializes when Office is ready
2. **Word.run()**: Creates a context for Word API operations
3. **context.sync()**: Executes queued operations
4. **Word APIs**: Access document elements (body, paragraphs, tables, etc.)

## Key Word API Features Used

- `context.document.body` - Access document content
- `context.document.getSelection()` - Get selected text
- `body.insertText()` - Insert text at specific locations
- `body.insertTable()` - Create tables
- Font formatting (bold, color, size)
- Table styling and formatting

## Troubleshooting

1. **SSL Certificate Issues**: 
   - First, open `https://localhost:3000/taskpane.html` in your browser and accept the certificate
   - Then the add-in should work in Word without SSL errors

2. **Certificate Issues After Windows Update**:
   If certificates become outdated after a Windows update:
   ```bash
   # Stop all Node.js processes first
   taskkill /f /im node.exe
   
   # Delete old certificates (replace {username} with your username)
   Remove-Item -Recurse -Force "C:\Users\{username}\.office-addin-dev-certs\*"
   
   # Or for current user:
   Remove-Item -Recurse -Force "$env:USERPROFILE\.office-addin-dev-certs\*"
   
   # Regenerate fresh certificates
   npx office-addin-dev-certs install
   ```

3. **Stop All Development Processes**:
   ```bash
   # Kill all Node.js processes
   taskkill /f /im node.exe
   
   # Kill Word processes if needed
   taskkill /f /im WINWORD.EXE
   
   # Check what's running on port 3000
   netstat -ano | findstr :3000
   
   # Check for remaining Node processes
   Get-Process | Where-Object {$_.ProcessName -like "*node*"}
   ```

4. **"Access Denied" loopback exemption error**: 
   - This warning is normal and doesn't prevent the add-in from working
   - To fix it completely, run PowerShell as Administrator and rerun `npm start`

5. **Add-in not loading**: 
   - Ensure both terminals are running (webpack dev server + Office debugging)
   - Verify the dev server is running on https://localhost:3000
   - Check that Word shows the "Word Integration" group in the Home tab

6. **Manifest errors**: Run `npm run validate` to check the manifest

7. **Word compatibility**: Requires Word 2016 or later, or Word Online

8. **Development Server Stops**: 
   - Keep Terminal 1 (webpack serve) running continuously
   - Only Terminal 2 (npm start) will exit after launching Word

## Next Steps

- Add more advanced Word features (content controls, styles, etc.)
- Integrate with external APIs
- Add authentication for cloud services
- Implement document templates
- Add custom ribbon commands