# Email Assistant Gmail Add-on

This folder contains the source code for the Gmail Add-on that connects to your backend for AI-generated email responses.

## Project Structure (Modular)

Apps Script supports multiple files—all share a global scope, so no imports are needed. Using `.js` files (configured via clasp `fileExtension`).

| File | Purpose |
|------|---------|
| **Config.js** | Script properties (SERVER_DOMAIN, FRONTEND_URL, etc.) |
| **Api.js** | Backend API calls (e.g. `getUserInfo`) |
| **Main.js** | Entry point – `buildAddOn` (homepage/contextual trigger) |
| **Handlers.js** | Action handlers: `onGenerateAIReply`, `onUseReply`, `onShowSettings`, `onUniversalSignOut` |

## Setup Instructions

1. **Install clasp (if not already):**
   ```sh
   npm install -g @google/clasp
   ```
2. **Login to clasp:**
   ```sh
   clasp login
   ```
3. **Push the code to Google Apps Script:**
   ```sh
   clasp push
   ```
4. **Open the project in the Apps Script Editor:**
   ```sh
   clasp open
   ```
5. **Test the Add-on:**
   - In the Apps Script Editor, click the "Deploy" button and select "Test deployments".
   - Authorize the add-on and install it for your Gmail account.
   - Open Gmail and you should see the add-on in the sidebar with a welcome message.

## Next Steps
- Implement backend communication and UI for AI response generation. 