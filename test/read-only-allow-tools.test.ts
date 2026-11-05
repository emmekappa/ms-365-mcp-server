import { afterEach, describe, expect, it, vi } from 'vitest';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn(), debug: vi.fn() },
}));

import { buildScopesFromEndpoints } from '../src/auth.js';
import { buildToolsRegistry } from '../src/graph-tools.js';
import { isDroppedByReadOnly, READ_ONLY_ALLOW_TOOLS_ENV } from '../src/lib/read-only.js';

const WRITE_TOOL = 'set-my-current-work-location';

afterEach(() => {
  delete process.env[READ_ONLY_ALLOW_TOOLS_ENV];
});

describe('isDroppedByReadOnly', () => {
  it('keeps everything when read-only mode is off', () => {
    expect(isDroppedByReadOnly(false, 'DELETE', 'delete-mail-message')).toBe(false);
  });

  it('keeps GETs and POST queries marked readOnly', () => {
    expect(isDroppedByReadOnly(true, 'get', 'list-mail-messages')).toBe(false);
    expect(isDroppedByReadOnly(true, 'post', 'get-schedule', { readOnly: true })).toBe(false);
  });

  it('drops write tools by default', () => {
    expect(isDroppedByReadOnly(true, 'post', WRITE_TOOL)).toBe(true);
    expect(isDroppedByReadOnly(true, 'patch', 'update-mail-message')).toBe(true);
  });

  it('keeps the write tools named by the allow list, and only those', () => {
    process.env[READ_ONLY_ALLOW_TOOLS_ENV] = WRITE_TOOL;
    expect(isDroppedByReadOnly(true, 'post', WRITE_TOOL)).toBe(false);
    expect(isDroppedByReadOnly(true, 'post', 'send-mail')).toBe(true);
  });

  it('matches the whole tool name', () => {
    process.env[READ_ONLY_ALLOW_TOOLS_ENV] = 'set-my';
    expect(isDroppedByReadOnly(true, 'post', WRITE_TOOL)).toBe(true);
  });
});

describe('work location tools', () => {
  it('are org-mode tools; read-only mode keeps the reads and drops the write', () => {
    const tools = [...buildToolsRegistry(true, true).keys()];
    expect(tools).toContain('get-my-work-hours-and-locations');
    expect(tools).toContain('get-my-work-location-occurrences');
    expect(tools).not.toContain(WRITE_TOOL);
    expect([...buildToolsRegistry(false, false).keys()]).not.toContain(
      'get-my-work-location-occurrences'
    );
  });

  it('read-only mode keeps the write when the allow list names it', () => {
    process.env[READ_ONLY_ALLOW_TOOLS_ENV] = WRITE_TOOL;
    const tools = [...buildToolsRegistry(true, true).keys()];
    expect(tools).toContain(WRITE_TOOL);
    expect(tools).not.toContain('create-calendar-event');
  });

  it('the allow list adds the scope of the write tool to the login scopes', () => {
    expect(buildScopesFromEndpoints(true, WRITE_TOOL, true)).toEqual([]);
    process.env[READ_ONLY_ALLOW_TOOLS_ENV] = WRITE_TOOL;
    expect(buildScopesFromEndpoints(true, WRITE_TOOL, true)).toEqual(['Calendars.ReadWrite']);
  });
});
