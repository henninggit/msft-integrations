# Word Integration Office Add-in

A Microsoft Word Office Add-in that demonstrates integration capabilities using Office.js APIs.

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

1. **Navigate to this directory**:
   ```bash
   cd integrations/word-addin
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Start the development environment** (requires 2 terminals):

   **Terminal 1 - Start Webpack Dev Server:**
   ```bash
   npm run dev
   ```
   This starts the dev server at `http://localhost:3000/` and must keep running.

   **Terminal 2 - Launch Word with Add-in:**
   ```bash
   npm start
   ```
   This will:
   - Install SSL certificates for local development
   - Launch Word with the add-in loaded
   - The add-in will appear in the **Home** tab under "Word Integration"

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

## Troubleshooting

See the main project README for detailed troubleshooting steps.
