import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerCreateReplyTool } from './createReply.js';
import { registerListFoldersTool } from './listFolders.js';
import { registerOrganizeMailTool } from './organizeMail.js';
import { registerReadMailTool } from './readMail.js';
import { registerSaveDraftTool } from './saveDraft.js';
import { registerSearchMailTool } from './searchMail.js';

export function registerAllTools(server: McpServer): void {
  registerOrganizeMailTool(server);
  registerReadMailTool(server);
  registerSearchMailTool(server);
  registerCreateReplyTool(server);
  registerSaveDraftTool(server);
  registerListFoldersTool(server);
}
