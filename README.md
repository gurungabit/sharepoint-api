# SharePoint Graph API (TypeScript/AWS Lambda)

This is a standalone, serverless Node.js API that acts as a secure bridge between Microsoft 365 (SharePoint/OneNote) and your AWS infrastructure. It is designed to be deployed as an AWS Lambda function behind Amazon API Gateway using Terraform.

## Prerequisites
1. Register an Application in Microsoft Entra ID (Azure AD).
2. Grant the application the following **Application Permissions** (Admin Consent Required):
   - `Sites.Read.All`
   - `Files.Read.All`
3. Generate a Client Secret.

## Environment Variables
The Lambda function requires the following environment variables to authenticate with Microsoft Graph:
- `TENANT_ID`: Your Entra ID Tenant ID.
- `CLIENT_ID`: The Application (Client) ID.
- `CLIENT_SECRET`: The generated Client Secret.

## Building for Terraform
Run the following command to bundle and minify the TypeScript code into a single file ready for Lambda deployment:
```bash
npm run build
```
The output will be located at `dist/index.js`. Point your Terraform `aws_lambda_function` resource to this file.

## Available Endpoints

### 1. List Sites
Returns a list of all SharePoint sites the application has access to.
```http
GET /sites
```

### 2. List Document Libraries (Drives)
Returns all document libraries (drives) within a specific SharePoint site.
```http
GET /sites/{siteId}/drives
```

### 3. List Files (Drive Items)
Returns the files and folders located in the root of a specific document library.
```http
GET /drives/{driveId}/root/children
```

### 4. Get File Download URL
Returns metadata and a secure, temporary `@microsoft.graph.downloadUrl` that can be used to download the actual file content (PDF, DOCX) without needing further authentication.
```http
GET /drives/{driveId}/items/{itemId}/content
```
**Response:**
```json
{
  "filename": "Product_Requirements.pdf",
  "downloadUrl": "https://tenant.sharepoint.com/temp-auth-link...",
  "webUrl": "https://tenant.sharepoint.com/...",
  "lastModifiedDateTime": "2026-02-27T10:00:00Z"
}
```

### 5. Global Search
Searches across all SharePoint sites and document libraries for a specific keyword.
```http
GET /search?q={query}
```
