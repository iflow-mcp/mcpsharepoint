import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { Client } from "@microsoft/microsoft-graph-client";
import { ConfidentialClientApplication } from "@azure/msal-node";
import dotenv from "dotenv";
dotenv.config();

// Initialize the MCP server and SharePoint connector
async function createSharepointMcpServer() {
  
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const driveId = process.env.DRIVE_ID;
  const siteId = process.env.SITE_ID;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET");
  }

  // Create the server
  const server = new McpServer({
    name: "SharePoint Server",
    version: "1.0.0"
  });

  // Initialize Microsoft Graph client
  const clientApp = new ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret,
      authority: `https://login.microsoftonline.com/${tenantId}`
    }
  });

  const graphClient = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const clientCredentialRequest = {
          scopes: ['https://graph.microsoft.com/.default'],
        };
        const response = await clientApp.acquireTokenByClientCredential(clientCredentialRequest);
        return response?.accessToken || '';
      }
    }
  });

  // Resource: Folder contents (root or specific folder)
  server.resource(
    "folder",
    new ResourceTemplate("sharepoint://folder/{folderId?}", { list: undefined }),
    async (uri) => {
      try {
        const driveUrl = siteId ? `/sites/${siteId}/drive` : '/me/drive';
        const items = await graphClient.api(`${driveUrl}/root/children`).get();
        return {
          contents: [{
            uri: uri.href,
            text: JSON.stringify(items, null, 2)
          }]
        };
      } catch (error) {
        return {
          contents: [{
            uri: uri.href,
            text: `Error fetching folder contents: ${error}`
          }]
        };
      }
    }
  );

  // Resource: Sites
  server.resource(
    "sites",
    "sharepoint://sites",
    async (uri) => {
      try {
        const sites = await graphClient.api('/sites').get();
        return {
          contents: [{
            uri: uri.href,
            text: JSON.stringify(sites, null, 2)
          }]
        };
      } catch (error) {
        return {
          contents: [{
            uri: uri.href,
            text: `Error fetching sites: ${error}`
          }]
        };
      }
    }
  );

  // Resource: Document content
  server.resource(
    "document",
    new ResourceTemplate("sharepoint://document/{documentId}", { list: undefined }),
    async (uri, { documentId }) => {
      try {
        const driveUrl = siteId ? `/sites/${siteId}/drive` : '/me/drive';
        const item = await graphClient.api(`${driveUrl}/items/${documentId}`).get();
        const content = await graphClient.api(`${driveUrl}/items/${documentId}/content`).get();
        return {
          contents: [{
            uri: uri.href,
            text: content || JSON.stringify(item, null, 2)
          }]
        };
      } catch (error) {
        return {
          contents: [{
            uri: uri.href,
            text: `Error fetching document: ${error}`
          }]
        };
      }
    }
  );

  // Tool: Search for documents
  server.tool(
    "search-documents",
    {
      query: z.string().describe("Search query to find documents"),
      maxResults: z.string().optional().describe("Maximum number of results to return (as a string)")
    },
    async ({ query, maxResults = "10" }) => {
      try {
        const driveUrl = siteId ? `/sites/${siteId}/drive` : '/me/drive';
        const driveItems = await graphClient.api(`${driveUrl}/root/children`).get();
        const results = driveItems.value.filter((item: any) =>
          item.name && item.name.toLowerCase().includes(query.toLowerCase())
        ).slice(0, parseInt(maxResults.toString(), 10));
        return {
          content: [{
            type: "text",
            text: JSON.stringify(results, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error searching documents: ${error}`
          }],
          isError: true
        };
      }
    }
  );

  // Download document content
  server.tool(
    "download-document",
    {
      documentId: z.string().describe("The ID of the document to download")
    },
    async ({ documentId }) => {
      try {
        const driveUrl = siteId ? `/sites/${siteId}/drive` : '/me/drive';
        const item = await graphClient.api(`${driveUrl}/items/${documentId}`).get();
        const content = await graphClient.api(`${driveUrl}/items/${documentId}/content`).get();
        return {
          content: [{
            type: "text",
            text: content || JSON.stringify(item, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error downloading document: ${error}`
          }],
          isError: true
        };
      }
    }
  );

  // Prompt: Search and summarize a document
  server.prompt(
    "document-summary",
    {
      documentId: z.string().describe("The ID of the document to summarize")
    },
    ({ documentId }) => ({
      messages: [{
        role: "user",
        content: {
          type: "text",
          text: `Please retrieve the document with ID ${documentId} using the sharepoint://document/${documentId} resource, then provide a concise summary of its key points, main topics, and important information.`
        }
      }]
    })
  );

  // Prompt: Find relevant documents
  server.prompt(
    "find-relevant-documents",
    {
      topic: z.string().describe("The topic or subject to find documents about"),
      maxResults: z.string().optional().describe("Maximum number of results to return (as a string)")
    },
    ({ topic, maxResults = "5" }) => ({
      messages: [{
        role: "user",
        content: {
          type: "text",
          text: `Please use the search-documents tool to find up to ${maxResults} documents related to "${topic}". For each document, provide the title, author, last modified date, and a brief description of what it appears to contain based on the metadata.`
        }
      }]
    })
  );

  // Prompt: Explore folder contents
  server.prompt(
    "explore-folder",
    {
      folderId: z.string().optional().describe("The ID of the folder to explore (leave empty for root folder)")
    },
    ({ folderId }) => ({
      messages: [{
        role: "user",
        content: {
          type: "text",
          text: folderId
            ? `Please explore the contents of the folder with ID ${folderId} using the sharepoint://folder/${folderId} resource. List all documents and subfolders, organizing them by type and providing key details about each item.`
            : `Please explore the contents of the root folder using the sharepoint://folder resource. List all documents and subfolders, organizing them by type and providing key details about each item.`
        }
      }]
    })
  );

  return server;
}

// Example usage
async function main() {

  // Create and start the server
  const server = await createSharepointMcpServer();

  // Connect using stdio transport
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

// Run the server
main().catch(error => {
  console.error("Error starting server:", error);
  process.exit(1);
});