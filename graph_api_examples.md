# MS Graph API Email Automation Guide

## Overview
This guide provides working examples for automating email processing using MS Graph API in Power Automate, specifically for copying emails between mailboxes and moving them to specific folders.

## Key Limitations
- **Cross-mailbox copy is NOT supported** - You cannot copy messages directly between different mailboxes
- **destinationId must be a folder ID**, not an email address

## Recommended Approach: Create New Message

### Step 1: Get Original Message Content
```http
GET https://graph.microsoft.com/v1.0/users/e-service@quillarrowlaw.com/messages/{messageId}
```

### Step 2: Create New Message in Destination User's Mailbox
```http
POST https://graph.microsoft.com/v1.0/users/lldistro@quillarrowlaw.com/messages
Content-Type: application/json
Authorization: Bearer {token}

{
  "subject": "{original subject}",
  "body": {
    "contentType": "html",
    "content": "{original body content}"
  },
  "toRecipients": [
    {
      "emailAddress": {
        "address": "lldistro@quillarrowlaw.com"
      }
    }
  ],
  "from": {
    "emailAddress": {
      "address": "e-service@quillarrowlaw.com"
    }
  },
  "isDraft": false
}
```

### Step 3: Move to Specific Folder
First, get the folder ID:
```http
GET https://graph.microsoft.com/v1.0/users/lldistro@quillarrowlaw.com/mailFolders
```

Then move the message:
```http
POST https://graph.microsoft.com/v1.0/users/lldistro@quillarrowlaw.com/messages/{newMessageId}/move
Content-Type: application/json
Authorization: Bearer {token}

{
  "destinationId": "{actual-folder-id}"
}
```

## Power Automate Configuration

### Required Permissions
Your Power Automate connection needs these Graph permissions:
- `Mail.ReadWrite.Shared` - To read from shared mailbox
- `Mail.ReadWrite` - To create messages in user mailboxes
- `Mail.Send` - If sending emails

### Authentication Setup
Use either:
1. **Service Account** with proper mailbox permissions
2. **Managed Identity** (recommended for security)

### Sample Power Automate Actions

#### Action 1: Get Original Message
- **Action**: HTTP
- **Method**: GET  
- **URI**: `https://graph.microsoft.com/v1.0/users/e-service@quillarrowlaw.com/messages/@{triggerOutputs()?['body/id']}`
- **Authentication**: Bearer token

#### Action 2: Create Message in User Mailbox
- **Action**: HTTP
- **Method**: POST
- **URI**: `https://graph.microsoft.com/v1.0/users/@{variables('targetUser')}/messages`
- **Body**: 
```json
{
  "subject": "@{outputs('GetOriginalMessage')?['body/subject']}",
  "body": {
    "contentType": "html", 
    "content": "@{outputs('GetOriginalMessage')?['body/body/content']}"
  },
  "toRecipients": [
    {
      "emailAddress": {
        "address": "@{variables('targetUser')}"
      }
    }
  ],
  "isDraft": false
}
```

#### Action 3: Move to Specific Folder
- **Action**: HTTP
- **Method**: POST
- **URI**: `https://graph.microsoft.com/v1.0/users/@{variables('targetUser')}/messages/@{outputs('CreateMessage')?['body/id']}/move`
- **Body**:
```json
{
  "destinationId": "@{variables('targetFolderId')}"
}
```

## Getting Folder IDs

### Method 1: List All Folders
```http
GET https://graph.microsoft.com/v1.0/users/{userEmail}/mailFolders
```

### Method 2: Use Well-Known Folder Names
Instead of folder IDs, you can use these well-known names:
- `inbox`
- `drafts`
- `sentitems`
- `deleteditems`
- `archive`
- `junkemail`

Example:
```json
{
  "destinationId": "archive"
}
```

## Handling Attachments

If you need to copy attachments, you'll need additional steps:

### Get Attachments from Original Message
```http
GET https://graph.microsoft.com/v1.0/users/e-service@quillarrowlaw.com/messages/{messageId}/attachments
```

### Add Attachments to New Message
```http
POST https://graph.microsoft.com/v1.0/users/{targetUser}/messages/{newMessageId}/attachments
Content-Type: application/json

{
  "@odata.type": "#microsoft.graph.fileAttachment",
  "name": "{attachment name}",
  "contentBytes": "{base64 encoded content}"
}
```

## Alternative: Forward Instead of Copy

If preserving the original message structure is important:

```http
POST https://graph.microsoft.com/v1.0/users/e-service@quillarrowlaw.com/messages/{messageId}/forward
Content-Type: application/json

{
  "toRecipients": [
    {
      "emailAddress": {
        "address": "lldistro@quillarrowlaw.com"
      }
    }
  ],
  "comment": "Automatically forwarded based on content analysis"
}
```

## Error Handling

Common errors and solutions:

### "ErrorInvalidIdMalformed"
- **Cause**: Using email address instead of folder ID
- **Solution**: Use actual folder ID or well-known folder name

### "ErrorAccessDenied"  
- **Cause**: Insufficient permissions
- **Solution**: Ensure proper Graph permissions are granted

### "ErrorItemNotFound"
- **Cause**: Invalid message ID or folder ID
- **Solution**: Verify IDs exist and are accessible

## Best Practices

1. **Use Managed Identity** for authentication when possible
2. **Cache folder IDs** to avoid repeated API calls
3. **Implement retry logic** for transient errors
4. **Use batch requests** for multiple operations
5. **Handle rate limiting** (429 errors) appropriately

## Testing with Graph Explorer

Test your API calls using Graph Explorer (https://developer.microsoft.com/graph/graph-explorer):

1. Sign in with appropriate permissions
2. Test GET operations first to verify access
3. Use POST operations to test message creation
4. Verify folder operations work as expected