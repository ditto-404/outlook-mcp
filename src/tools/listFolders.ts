import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { outlookClient } from '../outlook/client.js';
import type { FolderNode } from '../outlook/types.js';

function renderTree(node: FolderNode, depth = 0): string[] {
  const indent = '  '.repeat(depth);
  const line = `${indent}- ${node.name} (${node.itemCount}건, 안읽음 ${node.unreadCount}건)`;
  return [line, ...node.children.flatMap((child) => renderTree(child, depth + 1))];
}

export function registerListFoldersTool(server: McpServer): void {
  server.registerTool(
    'list_folders',
    {
      title: '폴더 목록 조회',
      description: '받은 편지함 하위의 폴더 구조와 각 폴더의 메일/안읽음 수를 조회합니다. 분류 결과 확인이나 디버깅에 사용하세요.',
      inputSchema: {},
    },
    async () => {
      const tree = await outlookClient.listFolders();
      return {
        content: [{ type: 'text', text: renderTree(tree).join('\n') }],
      };
    },
  );
}
