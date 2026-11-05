import { afterEach, describe, expect, it, vi } from 'vitest';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn(), debug: vi.fn() },
}));

import {
  COLLEAGUE_WORK_LOCATIONS_ENV,
  COLLEAGUE_WORK_LOCATIONS_TOOL,
  getColleagueWorkLocations,
  validateRequest,
  type ColleagueWorkLocationsDeps,
} from '../src/lib/colleague-work-locations.js';
import { UTILITY_TOOLS, isUtilityToolAvailable } from '../src/graph-tools.js';

const TENANT = '11111111-1111-1111-1111-111111111111';
const CLIENT = '22222222-2222-2222-2222-222222222222';
const GRAPH = 'https://graph.example';

function jwt(claims: Record<string, unknown>): string {
  const enc = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${enc({ alg: 'none' })}.${enc(claims)}.sig`;
}

const CALLER = jwt({ tid: TENANT, appid: CLIENT });
const ANNA_ID = '33333333-3333-3333-3333-333333333333';

const REQUEST = {
  users: ['anna@example.com'],
  startDateTime: '2026-11-06T00:00:00Z',
  endDateTime: '2026-11-06T23:59:59Z',
};

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status });
}

function deps(routes: (url: string, auth: string) => Response): {
  deps: ColleagueWorkLocationsDeps;
  calls: string[];
} {
  const calls: string[] = [];
  return {
    calls,
    deps: {
      secrets: { clientId: CLIENT, tenantId: TENANT, cloudType: 'global' },
      graphBaseUrl: GRAPH,
      getAppToken: async () => 'app-token',
      fetch: async (url, init) => {
        calls.push(url);
        const auth = (init?.headers as Record<string, string>).Authorization;
        return routes(url, auth);
      },
    },
  };
}

const OCCURRENCE = {
  id: 'secret-id',
  workLocationType: 'office',
  start: { dateTime: '2026-11-06T09:00:00', timeZone: 'UTC' },
  end: { dateTime: '2026-11-06T18:00:00', timeZone: 'UTC' },
  placeId: 'building-1',
  timeOffDetails: { subject: 'Doctor' },
};

function graph(level: string) {
  return (url: string, auth: string) => {
    if (url.endsWith('/v1.0/me?$select=id')) {
      return auth === `Bearer ${CALLER}` ? json({ id: 'me' }) : json({}, 401);
    }
    if (url.endsWith('/v1.0/users/anna%40example.com?$select=id')) {
      // Resolved with the caller's token, not the app's
      expect(auth).toBe(`Bearer ${CALLER}`);
      return json({ id: ANNA_ID });
    }
    expect(auth).toBe('Bearer app-token');
    // occurrencesView fails on a UPN, so the work plan is always read by object id
    expect(url).toContain(`/users/${ANNA_ID}/settings/`);
    if (url.includes('occurrencesView')) return json({ value: [OCCURRENCE] });
    if (url.includes('maxSharedWorkLocationDetails')) {
      return json({ maxSharedWorkLocationDetails: level });
    }
    return json({}, 404);
  };
}

afterEach(() => {
  delete process.env[COLLEAGUE_WORK_LOCATIONS_ENV];
});

describe('caller check', () => {
  it('refuses a request without a token, and never asks for the app token', async () => {
    const getAppToken = vi.fn();
    const { deps: d } = deps(graph('approximate'));
    await expect(
      getColleagueWorkLocations(undefined, REQUEST, { ...d, getAppToken })
    ).rejects.toThrow(/No signed-in user/);
    expect(getAppToken).not.toHaveBeenCalled();
  });

  it('refuses a token of another tenant or another app', async () => {
    const { deps: d } = deps(graph('approximate'));
    await expect(
      getColleagueWorkLocations(jwt({ tid: 'other', appid: CLIENT }), REQUEST, d)
    ).rejects.toThrow(/not issued to this app/);
    await expect(
      getColleagueWorkLocations(jwt({ tid: TENANT, appid: 'other' }), REQUEST, d)
    ).rejects.toThrow(/not issued to this app/);
  });

  it('refuses a token Graph rejects', async () => {
    const forged = jwt({ tid: TENANT, appid: CLIENT, forged: true });
    const { deps: d } = deps(graph('approximate'));
    await expect(getColleagueWorkLocations(forged, REQUEST, d)).rejects.toThrow(/HTTP 401/);
  });
});

describe('sharing level', () => {
  it('approximate: type and hours only, never ids, place or time-off details', async () => {
    const { deps: d } = deps(graph('approximate'));
    const [result] = await getColleagueWorkLocations(CALLER, REQUEST, d);
    expect(result).toEqual({
      user: 'anna@example.com',
      shared: true,
      detail: 'approximate',
      occurrences: [{ workLocationType: 'office', start: OCCURRENCE.start, end: OCCURRENCE.end }],
    });
  });

  it('specific: adds the place id', async () => {
    const { deps: d } = deps(graph('specific'));
    const [result] = await getColleagueWorkLocations(CALLER, REQUEST, d);
    expect(result).toMatchObject({ detail: 'specific', occurrences: [{ placeId: 'building-1' }] });
  });

  it('none: returns nothing and does not read the work plan', async () => {
    const { deps: d, calls } = deps(graph('none'));
    const [result] = await getColleagueWorkLocations(CALLER, REQUEST, d);
    expect(result).toEqual({
      user: 'anna@example.com',
      shared: false,
      reason: 'The user does not share their work location.',
    });
    expect(calls.some((c) => c.includes('occurrencesView'))).toBe(false);
  });

  it('an unreadable sharing level returns an error, not the work plan', async () => {
    const { deps: d, calls } = deps((url, auth) =>
      url.includes('maxSharedWorkLocationDetails') ? json({}, 403) : graph('specific')(url, auth)
    );
    const [result] = await getColleagueWorkLocations(CALLER, REQUEST, d);
    expect(result).toHaveProperty('error');
    expect(calls.some((c) => c.includes('occurrencesView'))).toBe(false);
  });
});

describe('user resolution and errors', () => {
  it('uses an object id as is, without a lookup', async () => {
    const { deps: d, calls } = deps(graph('approximate'));
    const [result] = await getColleagueWorkLocations(CALLER, { ...REQUEST, users: [ANNA_ID] }, d);
    expect(result).toMatchObject({ user: ANNA_ID, shared: true });
    expect(calls.some((c) => c.includes('/users/') && c.endsWith('?$select=id'))).toBe(false);
  });

  it('reports a user the directory does not know, with the Graph message', async () => {
    const { deps: d } = deps((url, auth) =>
      url.includes('/users/') && url.endsWith('?$select=id')
        ? json({ error: { message: "Resource 'x' does not exist" } }, 404)
        : graph('approximate')(url, auth)
    );
    const [result] = await getColleagueWorkLocations(CALLER, REQUEST, d);
    expect(result).toEqual({
      user: 'anna@example.com',
      error: "User not found. Resource 'x' does not exist",
    });
  });
});

describe('validateRequest', () => {
  it('accepts a normal request', () => {
    expect(validateRequest(REQUEST)).toBeUndefined();
  });

  it('refuses path characters, long ranges and reversed ranges', () => {
    expect(validateRequest({ ...REQUEST, users: ['a/../../me'] })).toBeDefined();
    expect(validateRequest({ ...REQUEST, users: ["x')/foo"] })).toBeDefined();
    expect(validateRequest({ ...REQUEST, endDateTime: '2027-01-30T00:00:00Z' })).toMatch(/31 days/);
    expect(validateRequest({ ...REQUEST, endDateTime: '2026-11-05T00:00:00Z' })).toMatch(/after/);
  });
});

describe('registration', () => {
  const tool = UTILITY_TOOLS.find((t) => t.name === COLLEAGUE_WORK_LOCATIONS_TOOL)!;
  const options = { readOnly: true, httpMode: true, scopeMode: 'org' as const };

  it('is a read-only tool, off unless the operator turns it on', () => {
    expect(isUtilityToolAvailable(tool, options)).toBe(false);
    process.env[COLLEAGUE_WORK_LOCATIONS_ENV] = 'true';
    expect(isUtilityToolAvailable(tool, options)).toBe(true);
  });
});
