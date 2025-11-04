#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
  ListToolsRequestSchema,
  CallToolRequestSchema,
  ListPromptsRequestSchema,
  GetPromptRequestSchema
} from "@modelcontextprotocol/sdk/types.js";
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
  const server = new Server({
    name: "SharePoint Server",
    version: "1.0.1"
  }, {
    capabilities: {
      resources: {},
      tools: {},
      prompts: {}
    }
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

  // Resource handlers
  server.setRequestHandler(ListResourcesRequestSchema, async () => {
    return {
      resources: [
        {
          uri: "sharepoint://folder",
          name: "SharePoint Folder",
          description: "Access SharePoint folder contents",
          mimeType: "application/json"
        },
        {
          uri: "sharepoint://sites",
          name: "SharePoint Sites",
          description: "List SharePoint sites",
          mimeType: "application/json"
        }
      ]
    };
  });

  server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
    const uri = request.params.uri;

    if (uri === "sharepoint://folder" || uri.startsWith("sharepoint://folder/")) {
      try {
        const driveUrl = siteId ? `/sites/${siteId}/drive` : '/me/drive';
        const items = await graphClient.api(`${driveUrl}/root/children`).get();
        return {
          contents: [{
            uri: uri,
            text: JSON.stringify(items, null, 2)
          }]
        };
      } catch (error) {
        return {
          contents: [{
            uri: uri,
            text: `Error fetching folder contents: ${error}`
          }]
        };
      }
    }

    if (uri === "sharepoint://sites") {
      try {
        const sites = await graphClient.api('/sites').get();
        return {
          contents: [{
            uri: uri,
            text: JSON.stringify(sites, null, 2)
          }]
        };
      } catch (error) {
        return {
          contents: [{
            uri: uri,
            text: `Error fetching sites: ${error}`
          }]
        };
      }
    }

    if (uri.startsWith("sharepoint://document/")) {
      try {
        const documentId = uri.replace("sharepoint://document/", "");
        const driveUrl = siteId ? `/sites/${siteId}/drive` : '/me/drive';
        const item = await graphClient.api(`${driveUrl}/items/${documentId}`).get();
        const content = await graphClient.api(`${driveUrl}/items/${documentId}/content`).get();
        return {
          contents: [{
            uri: uri,
            text: content || JSON.stringify(item, null, 2)
          }]
        };
      } catch (error) {
        return {
          contents: [{
            uri: uri,
            text: `Error fetching document: ${error}`
          }]
        };
      }
    }

    throw new Error(`Unsupported URI: ${uri}`);
  });

  // Tool handlers
  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: [
        {
          name: "search-documents",
          description: "Search for documents in SharePoint",
          inputSchema: {
            type: "object",
            properties: {
              query: {
                type: "string",
                description: "Search query to find documents"
              },
              maxResults: {
                type: "string",
                description: "Maximum number of results to return (as a string)",
                default: "10"
              }
            },
            required: ["query"]
          }
        }
      ]
    };
  });

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    if (name === "search-documents") {
      const { query, maxResults = "10" } = args as { query: string; maxResults?: string };
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

    throw new Error(`Unknown tool: ${name}`);
  });

  // Prompt handlers
  server.setRequestHandler(ListPromptsRequestSchema, async () => {
    return {
      prompts: [
        {
          name: "document-summary",
          description: "Get a summary of a SharePoint document",
          arguments: [
            {
              name: "documentId",
              description: "The ID of the document to summarize",
              required: true
            }
          ]
        },
        {
          name: "find-relevant-documents",
          description: "Find documents relevant to a topic",
          arguments: [
            {
              name: "topic",
              description: "The topic or subject to find documents about",
              required: true
            },
            {
              name: "maxResults",
              description: "Maximum number of results to return (as a string)",
              required: false
            }
          ]
        }
      ]
    };
  });

  server.setRequestHandler(GetPromptRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    if (name === "document-summary") {
      const { documentId } = args as { documentId: string };
      return {
        messages: [{
          role: "user" as const,
          content: {
            type: "text" as const,
            text: `Please retrieve the document with ID ${documentId} using the sharepoint://document/${documentId} resource, then provide a concise summary of its key points, main topics, and important information.`
          }
        }]
      };
    }

    if (name === "find-relevant-documents") {
      const { topic, maxResults = "5" } = args as { topic: string; maxResults?: string };
      return {
        messages: [{
          role: "user" as const,
          content: {
            type: "text" as const,
            text: `Please use the search-documents tool to find up to ${maxResults} documents related to "${topic}". For each document, provide the title, author, last modified date, and a brief description of what it appears to contain based on the metadata.`
          }
        }]
      };
    }

    throw new Error(`Unknown prompt: ${name}`);
  });

  return server;
}

// Main function
async function main() {
  try {
    // Create and start the server
    const server = await createSharepointMcpServer();

    // Connect using stdio transport
    const transport = new StdioServerTransport();
    await server.connect(transport);
  } catch (error) {
    console.error("Error starting server:", error);
    process.exit(1);
  }
}

// Run the server
main();