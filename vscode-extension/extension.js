const vscode = require('vscode');

const PROVIDER_ID = 'word-mcp';
const PACKAGE_ID = 'WordMcp.McpServer';

/**
 * Registers the Word MCP server with VS Code.
 *
 * The server itself is not bundled. It is a .NET tool published to nuget.org and launched through
 * `dnx`, which resolves the package on demand - that keeps the extension a few kilobytes and lets
 * the server be updated without shipping a new extension.
 *
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
  const changed = new vscode.EventEmitter();
  context.subscriptions.push(changed);

  context.subscriptions.push(
    vscode.lm.registerMcpServerDefinitionProvider(PROVIDER_ID, {
      onDidChangeMcpServerDefinitions: changed.event,
      provideMcpServerDefinitions: () => {
        // Word automation is Windows-only, so offering the server anywhere else would only produce
        // a server that fails on first use.
        if (process.platform !== 'win32') {
          return [];
        }

        const settings = vscode.workspace.getConfiguration('wordMcp');
        const version = (settings.get('serverVersion') || '').trim();
        const spec = version ? `${PACKAGE_ID}@${version}` : PACKAGE_ID;

        const env = {};
        if (settings.get('serviceMode') === 'daemon') {
          env.WORDMCP_SERVICE_MODE = 'daemon';
        }

        return [
          new vscode.McpStdioServerDefinition(
            'Word',
            'dnx',
            [spec, '--yes'],
            env,
            version || undefined
          )
        ];
      }
    })
  );

  context.subscriptions.push(
    vscode.workspace.onDidChangeConfiguration((event) => {
      if (event.affectsConfiguration('wordMcp')) {
        changed.fire();
      }
    })
  );
}

function deactivate() {}

module.exports = { activate, deactivate };
