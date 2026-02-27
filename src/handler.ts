import type { APIGatewayProxyEvent, APIGatewayProxyResult } from "aws-lambda";
import { getGraphClient } from "./graph.js";

export const handler = async (
  event: APIGatewayProxyEvent,
): Promise<APIGatewayProxyResult> => {
  try {
    const client = getGraphClient();
    const path = event.path;
    const httpMethod = event.httpMethod;

    console.log(`Received request: ${httpMethod} ${path}`);

    // Route: GET /sites
    // Lists all SharePoint sites the app has access to
    if (httpMethod === "GET" && path === "/sites") {
      const sites = await client.api("/sites?search=*").get();
      return success(sites);
    }

    // Route: GET /sites/{siteId}/drives
    // Get document libraries (drives) for a specific site
    const siteDrivesMatch = path.match(/^\/sites\/([^/]+)\/drives$/);
    if (httpMethod === "GET" && siteDrivesMatch) {
      const siteId = siteDrivesMatch[1];
      const drives = await client.api(`/sites/${siteId}/drives`).get();
      return success(drives);
    }

    // Route: GET /drives/{driveId}/root/children
    // List files and folders in the root of a specific document library
    const driveRootMatch = path.match(/^\/drives\/([^/]+)\/root\/children$/);
    if (httpMethod === "GET" && driveRootMatch) {
      const driveId = driveRootMatch[1];
      const children = await client
        .api(`/drives/${driveId}/root/children`)
        .get();
      return success(children);
    }

    // Route: GET /drives/{driveId}/items/{itemId}/content
    // Returns the temporary, unauthenticated download URL for a file
    const itemContentMatch = path.match(
      /^\/drives\/([^/]+)\/items\/([^/]+)\/content$/,
    );
    if (httpMethod === "GET" && itemContentMatch) {
      const driveId = itemContentMatch[1];
      const itemId = itemContentMatch[2];

      const itemMeta = await client
        .api(`/drives/${driveId}/items/${itemId}`)
        .get();
      const downloadUrl = itemMeta["@microsoft.graph.downloadUrl"];

      if (downloadUrl) {
        return success({
          filename: itemMeta.name,
          downloadUrl: downloadUrl,
          webUrl: itemMeta.webUrl,
          lastModifiedDateTime: itemMeta.lastModifiedDateTime,
        });
      } else {
        return error(
          "Download URL not found for this item. It might be a folder.",
          404,
        );
      }
    }

    // Route: GET /search?q={query}
    // Search across all SharePoint sites and OneDrive using Graph Search API
    if (httpMethod === "GET" && path === "/search") {
      const query = event.queryStringParameters?.q;
      if (!query) return error("Missing search query parameter 'q'", 400);

      const searchParams = {
        requests: [
          {
            entityTypes: ["listItem", "driveItem"],
            query: { queryString: query },
          },
        ],
      };
      const results = await client.api("/search/query").post(searchParams);
      return success(results);
    }

    // Fallback Route
    return error(`Route not found: ${httpMethod} ${path}`, 404);
  } catch (err: any) {
    console.error("Error processing request:", err);
    return error(err.message || "Internal Server Error", err.statusCode || 500);
  }
};

const success = (data: any): APIGatewayProxyResult => ({
  statusCode: 200,
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(data),
});

const error = (message: string, statusCode = 500): APIGatewayProxyResult => ({
  statusCode,
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({ error: message }),
});
