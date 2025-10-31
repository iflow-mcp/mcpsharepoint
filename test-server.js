import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

// Create a simple test server without authentication
async function createTestServer() {
  const server = new McpServer({
    name: "SharePoint Test Server",
    version: "1.0.0"
  });

  // Simple test tool
  server.tool(
    "test-connection",
    {
      message: z.string().optional().describe("Test message")
    },
    async ({ message = "Hello SharePoint!" }) => {
      return {
        content: [{
          type: "text",
          text: `SharePoint MCP Server is working! Message: ${message}`
        }]
      };
    }
  );

  return server;
}

async function main() {
  const server = await createTestServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch(error => {
  console.error("Error starting test server:", error);
  process.exit(1);
});