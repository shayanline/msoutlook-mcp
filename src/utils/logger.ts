/**
 * Minimal logger that writes to stderr so MCP stdout stays clean.
 */

const isDebug = process.env.MSOUTLOOK_DEBUG === 'true';

export const logger = {
  info: (msg: string, ...args: unknown[]) => {
    process.stderr.write(`[msoutlook-mcp] ${msg}${args.length ? ' ' + JSON.stringify(args) : ''}\n`);
  },
  debug: (msg: string, ...args: unknown[]) => {
    if (!isDebug) return;
    process.stderr.write(`[msoutlook-mcp:debug] ${msg}${args.length ? ' ' + JSON.stringify(args) : ''}\n`);
  },
  warn: (msg: string, ...args: unknown[]) => {
    process.stderr.write(`[msoutlook-mcp:warn] ${msg}${args.length ? ' ' + JSON.stringify(args) : ''}\n`);
  },
  error: (msg: string, ...args: unknown[]) => {
    process.stderr.write(`[msoutlook-mcp:error] ${msg}${args.length ? ' ' + JSON.stringify(args) : ''}\n`);
  },
};
