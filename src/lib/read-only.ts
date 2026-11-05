/**
 * Read-only mode keeps a tool when it reads: a GET, or a POST that endpoints.json
 * marks `readOnly` (a query sent as POST, such as get-schedule).
 *
 * MS365_MCP_READ_ONLY_ALLOW_TOOLS names write tools that read-only mode keeps
 * anyway: a case-insensitive regex matched against the whole tool name. It lets a
 * deployment stay read-only and still offer a narrow write, such as
 * set-my-current-work-location, without turning on every other write tool.
 * The tools it lets through stay write tools: their annotations do not change.
 */
export const READ_ONLY_ALLOW_TOOLS_ENV = 'MS365_MCP_READ_ONLY_ALLOW_TOOLS';

interface ReadOnlyEndpoint {
  toolName?: string;
  readOnly?: boolean;
}

let cachedPattern: string | undefined;
let cachedRegex: RegExp | undefined;

/**
 * Compile the allow pattern. Anchored, so `set-my-current-work-location` never
 * also matches a longer tool name. Throws on an invalid regex.
 */
export function compileReadOnlyAllowTools(pattern: string): RegExp {
  return new RegExp(`^(?:${pattern})$`, 'i');
}

function readOnlyAllowRegex(): RegExp | undefined {
  const pattern = process.env[READ_ONLY_ALLOW_TOOLS_ENV]?.trim() || undefined;
  if (pattern !== cachedPattern) {
    cachedPattern = pattern;
    cachedRegex = pattern ? compileReadOnlyAllowTools(pattern) : undefined;
  }
  return cachedRegex;
}

/** True when read-only mode drops this tool. */
export function isDroppedByReadOnly(
  readOnly: boolean,
  method: string,
  toolName: string,
  endpoint?: ReadOnlyEndpoint
): boolean {
  if (!readOnly) return false;
  const verb = method.toUpperCase();
  if (verb === 'GET') return false;
  if (verb === 'POST' && endpoint?.readOnly) return false;
  return !readOnlyAllowRegex()?.test(toolName);
}
