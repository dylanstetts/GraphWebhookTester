# Microsoft Graph Security Webhook Tester

A comprehensive Python GUI application for testing Microsoft Graph change notifications with **enhanced change analysis and operation detection**, specifically designed to help diagnose issues with security-related user action webhooks (file share, "Copy link," folder actions, etc.).

## Latest Features (October 2025)

- **Enhanced Operation Detection**: Specific operation type identification (permission grants, file uploads, renames, moves, etc.)
- **Smart Correlation Analysis**: Advanced timing-based correlation between webhooks and actual file changes
- **Security-Focused Analysis**: Priority scoring for permission changes and security-related operations
- **Professional Interface**: Clean, emoji-free interface suitable for professional presentations
- **Organized Project Structure**: Automatic file organization with dedicated folders for logs, webhooks, and analysis
- **Subscription Management**: View, refresh, and delete subscriptions directly from the GUI
- **Real-time Analysis**: Analyze webhook notifications as they arrive with comprehensive change tracking
- **Operation Details**: Show exactly who got access to what files and when
- **Webhook Correlation**: Accurate timing analysis showing webhook latency and correlation confidence

## Core Features

### Enhanced Change Analysis Engine
- **Operation Type Detection**: Distinguishes between permission_granted, file_uploaded, file_renamed, file_moved, file_deleted, etc.
- **User Activity Tracking**: Identifies who performed actions with specific user details
- **Timing Correlation**: Correlates webhook notifications with actual Microsoft Graph activities
- **Permission Analysis**: Detailed analysis of sharing and permission changes
- **Security Event Focus**: Prioritizes security-related activities in analysis

### Modern GUI Interface
- **Tabbed Interface**: Intuitive organization of functionality
- **Dual Authentication**: Both interactive (user) and app-only (client credentials) authentication
- **Security Webhooks**: Full support for `Prefer: includesecuritywebhooks` header
- **Audio Notifications**: Sound feedback for webhook creation and errors
- **Real-time Monitoring**: Live subscription status and management

### File Organization
- **Automatic Folder Structure**: Organized storage of all application data
- **Comprehensive Logging**: Detailed logging system with separate files for different components
- **Change History**: Browse and analyze historical webhook data
- **Analysis Archive**: Complete archive of all analysis results

## Project Structure

The application automatically organizes files into structured folders:

```
SecurityWebhooks - ODSP/
├── logs/                           # All log files
│   ├── enhanced_changes.log        # Enhanced tracker logs
│   ├── graph_api_requests.log      # HTTP request/response logs
│   └── delta_changes.log           # Legacy delta tracking logs
├── webhook_notifications/          # Received webhook files
│   └── webhook_notification_*.json # Individual webhook notifications
├── change_analysis/               # Analysis results
│   ├── enhanced_analysis_*.json   # Detailed change analysis files
│   └── change_details_*.json      # Change detail summaries
├── config.json                    # Your configuration
├── graph_security_webhook_tester.py # Main application
├── enhanced_change_tracker.py     # Enhanced analysis engine
├── webhook_receiver.py            # Local webhook receiver
└── README.md                      # This file
```

## Prerequisites

1. **Microsoft 365 Developer Account** or access to a Microsoft 365 tenant
2. **Azure App Registration** with appropriate permissions
3. **Python 3.8+** installed on your system
4. **ngrok** (optional but recommended for webhook testing)

## Azure App Registration Setup

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Configure your app:
   - **Name**: Graph Security Webhook Tester
   - **Supported account types**: Choose based on your needs
   - **Redirect URI**: 
     - For interactive auth: `http://localhost` (Public client/native)
     - For web apps: Your actual redirect URI

4. **API Permissions** (Microsoft Graph):
   - `Files.ReadWrite.All` (Application/Delegated)
   - `Sites.ReadWrite.All` (Application/Delegated)
   - `User.Read.All` (Application/Delegated) - for enhanced analysis
   - `AuditLog.Read.All` (Application) - for security audit data

5. **Grant admin consent** for the permissions

6. **Client Secret** (if using app-only authentication):
   - Go to "Certificates & secrets"
   - Create a new client secret
   - Copy the secret value (you won't see it again!)

## Installation & Setup

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure the Application
```bash
cp config_template.json config.json
```

Edit `config.json` with your app registration details:
```json
{
  "client_id": "your-app-client-id-here",
  "client_secret": "your-app-client-secret-here-optional",
  "tenant_id": "common",
  "auth_type": "interactive",
  "subscription_defaults": {
    "resource": "/me/drive/root",
    "change_type": "updated",
    "notification_url": "http://localhost:8000",
    "expiration_hours": "24",
    "include_security_webhooks": true
  }
}
```

### 3. Set Up Webhook Receiving

#### Option A: Using ngrok (Recommended)
1. **Install ngrok**: Download from [ngrok.com](https://ngrok.com)
2. **Run the webhook receiver**:
   ```bash
   python webhook_receiver.py
   ```
3. **In a new terminal, start ngrok**:
   ```bash
   ngrok http 8000
   ```
4. **Copy the ngrok URL** (e.g., `https://abc123.ngrok-free.app`) and use it as your notification URL

#### Option B: Using webhook.site
1. Go to [webhook.site](https://webhook.site)
2. Copy your unique URL
3. Use it as the notification URL in the app

## Quick Start Guide

### 1. Launch the Application
```bash
python graph_security_webhook_tester.py
```

### 2. Authentication
- Go to the **Authentication** tab
- Your app details should auto-load from `config.json`
- Click **"Authenticate"**
- Complete the authentication flow

### 3. Create a Subscription
- Go to the **Create Subscription** tab
- Configure your subscription:
  - **Resource**: `/me/drive/root` (auto-loaded)
  - **Change Type**: `updated` (auto-loaded)
  - **Notification URL**: Your ngrok or webhook.site URL
  - **Expiration**: `24` hours (auto-loaded)
  - **Include Security Webhooks**: Enabled (auto-loaded)
- Click **"Create Subscription"**

### 4. Start Webhook Receiver
```bash
python webhook_receiver.py
```
The receiver will save all webhook notifications to the `webhook_notifications/` folder.

### 5. Test Security Actions
Perform security-related actions:
- Share a file or folder
- Create sharing links ("Copy link")
- Change file/folder permissions
- Add/remove users from shared content

### 6. Analyze Results
- Go to the **Change Analysis** tab
- Click **"Analyze Latest Webhook"** for automatic analysis
- Or select a specific webhook file and click **"Analyze Selected"**
- Browse **"Refresh Changes"** to see all historical analysis

## Enhanced Analysis Features

### Operation Type Detection
The enhanced change tracker identifies specific operations:
- **permission_granted**: Access granted to specific users (e.g., "Access granted to: Holly Holt")
- **file_uploaded**: New file uploads to monitored folders
- **file_renamed**: File renaming operations with old filename details
- **file_moved**: File moves with source location information
- **file_content_modified**: Content changes with version information
- **file_deleted**: File deletion operations
- **file_restored**: File restoration from deletion
- **file_copied**: File copy operations

### Correlation Analysis
- **Timing Correlation**: Matches webhook notifications with actual Graph API activities
- **Latency Measurement**: Shows exact time between action and webhook notification
- **Confidence Scoring**: Rates correlation confidence (high/medium/low)
- **Priority Scoring**: Prioritizes security-related operations in analysis

### Analysis Output Example
```
CORRELATION ANALYSIS:
  Matched Item: TeamsUserActivity_Report_2025-06-30.csv
  Change Type: permission_granted
  Operation: Access granted to: Holly Holt
  Change Time: 2025-01-20 13:35:00
  Webhook Time: 2025-01-20T18:35:35
  Latency: 35.0 seconds
  Confidence: high
```

### Automatic Analysis Workflow
1. **Webhook Reception**: Notifications saved to `webhook_notifications/`
2. **Automatic Enhancement**: Deep Graph API analysis triggered
3. **Operation Detection**: Specific operation types and details identified
4. **Correlation Analysis**: Timing-based correlation with confidence scoring
5. **Detailed Reports**: Comprehensive analysis saved to `change_analysis/`
6. **GUI Integration**: View results directly in the application

## Subscription Management

### View Subscriptions
- **Monitor Subscriptions** tab shows all active subscriptions
- Real-time status updates
- Subscription details and expiration times

### Manage Subscriptions
- **Refresh Subscriptions**: Update the subscription list
- **Refresh List**: Update the dropdown selection
- **Delete Subscription**: Remove subscriptions directly from the GUI

## Common Resources for Testing

| Resource | Description | Use Case |
|----------|-------------|----------|
| `/me/drive/root` | User's OneDrive root | Personal file sharing |
| `/sites/{site-id}/drive/root` | SharePoint site drive | Team file sharing |
| `/me/drive/items/{item-id}` | Specific file/folder | Targeted monitoring |
| `/groups/{group-id}/drive/root` | Microsoft 365 Group drive | Group collaboration |

## Troubleshooting

### Port Issues
**Problem**: Port 8000 already in use
```bash
# Check what's using the port
netstat -ano | findstr :8000

# Kill the process (Windows)
taskkill /PID <PID> /F

# Change port in webhook_receiver.py if needed
```

### Authentication Issues
- **Error**: "AADSTS65001: The user or administrator has not consented to use the application"
  - **Solution**: Ensure admin consent is granted for all required permissions
  
- **Error**: "AADSTS50011: No reply address is registered for the application"
  - **Solution**: Add `http://localhost` as a redirect URI in your app registration

### Subscription Issues
- **Error**: "Subscription validation request failed"
  - **Solution**: Ensure your notification URL is publicly accessible and returns a 200 OK with the validation token

- **Error**: "Insufficient privileges to complete the operation"
  - **Solution**: Check that your app has the required permissions and admin consent

### Enhanced Analysis Issues
- **No analysis results**: Verify the webhook notifications are being saved to `webhook_notifications/`
- **API errors**: Check the `logs/graph_api_requests.log` for detailed error information
- **Missing permissions**: Ensure all required Graph API permissions are granted

## File Management

### Automatic Organization
- **Logs**: All logging output in `logs/` folder
- **Webhooks**: Received notifications in `webhook_notifications/` folder  
- **Analysis**: Enhanced analysis results in `change_analysis/` folder

### File Naming Conventions
- **Webhooks**: `webhook_notification_YYYYMMDD_HHMMSS.json`
- **Analysis**: `enhanced_analysis_YYYYMMDD_HHMMSS.json`
- **Logs**: Timestamped entries in respective log files

## Audio Notifications

The application provides audio feedback:
- **Success Sound**: When subscriptions are created successfully
- **Error Sound**: When errors occur during operations

## Security Considerations

- **Never commit** your `config.json` file with real credentials
- **Use client secrets** securely and rotate them regularly
- **Limit permissions** to only what's necessary for your testing
- **Use HTTPS** for all webhook endpoints in production
- **Monitor logs** for any suspicious activity

## Log Files

- **API Logs**: `logs/graph_api_requests.log` - All HTTP requests and responses
- **Enhanced Logs**: `logs/enhanced_changes.log` - Enhanced analysis tracking
- **Application Logs**: Console output for application-level events

## Testing Workflow

### Complete Testing Process
1. **Setup**: Configure app registration and authentication
2. **Start Services**: Launch webhook receiver and ngrok
3. **Create Subscription**: Set up monitoring for your target resource
4. **Trigger Actions**: Perform security-related file operations
5. **Monitor Reception**: Check webhook_notifications/ folder for incoming notifications
6. **Analyze Results**: Use the enhanced analysis features to understand changes
7. **Review Details**: Browse historical analysis in the Change Analysis tab

### Best Practices
- **Start with simple resources** like `/me/drive/root`
- **Test incremental changes** (one action at a time)
- **Monitor both webhook reception and analysis**
- **Keep ngrok running** throughout your testing session
- **Review logs regularly** for troubleshooting

## Advanced Features

### Operation Priority Scoring
The analysis engine uses sophisticated scoring to prioritize different types of operations:
- **Permission Changes**: Highest priority (60% score multiplier)
- **File Operations**: High priority (80% multiplier for uploads, renames, moves)
- **Content Modifications**: Medium priority (90% multiplier)

### Timezone Handling
- **Automatic Conversion**: Handles EST/EDT timezone conversions
- **Accurate Latency**: Precise webhook timing analysis
- **Correlation Windows**: Configurable time windows for correlation analysis

### Professional Presentation
- **Structured Output**: Organized analysis results suitable for reporting
- **Detailed Logging**: Comprehensive logging for audit and troubleshooting

## Legal Notice

This tool is for testing and diagnostic purposes only. Ensure you comply with your organization's policies and Microsoft's terms of service when testing with production data.

---

**Version**: Latest (October 2025)  
**Compatibility**: Python 3.8+, Microsoft Graph v1.0  
**License**: Use in accordance with your organization's policies