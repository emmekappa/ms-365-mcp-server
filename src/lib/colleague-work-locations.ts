/**
 * Colleagues' work locations (office, remote, time off) for a date range.
 *
 * Microsoft Graph lets delegated callers read only their own work plan
 * (/users/{id}/settings/workHoursAndLocations needs the signed-in user's id). Outlook,
 * Teams and Places show colleagues' plans, but Graph exposes them only to application
 * permissions. This tool therefore calls Graph with the app's own token
 * (client credentials), which can read every calendar the app is granted. The
 * operator limits that grant to a group of mailboxes with RBAC for Applications in
 * Exchange Online (role "Application Calendars.Read" and a management scope).
 *
 * The server never returns more than the colleague shares in Outlook:
 * - it reads only work plan occurrences, never events;
 * - it honours maxSharedWorkLocationDetails: `none` hides the user, `specific` adds
 *   the place id, anything else returns only the location type and hours;
 * - it never returns time-off details or occurrence ids.
 *
 * The server does not validate bearer tokens itself (Graph does), so before using the
 * app token it checks that the caller holds a token Graph accepts, issued to this app
 * in this tenant. Off unless MS365_MCP_COLLEAGUE_WORK_LOCATIONS=true.
 */
import { AppTokenClient } from '../app-mode.js';
import { getGraphBaseUrl } from '../cloud-config.js';
import { getSecrets, type AppSecrets } from '../secrets.js';

export const COLLEAGUE_WORK_LOCATIONS_ENV = 'MS365_MCP_COLLEAGUE_WORK_LOCATIONS';
export const COLLEAGUE_WORK_LOCATIONS_TOOL = 'get-colleague-work-locations';
export const MAX_USERS = 25;
export const MAX_RANGE_DAYS = 31;

export function colleagueWorkLocationsEnabled(): boolean {
  return ['true', '1', 'yes'].includes(
    (process.env[COLLEAGUE_WORK_LOCATIONS_ENV] ?? '').trim().toLowerCase()
  );
}

export type FetchLike = (
  url: string,
  init?: { headers?: Record<string, string> }
) => Promise<Response>;

export interface ColleagueWorkLocationsDeps {
  secrets: AppSecrets;
  graphBaseUrl: string;
  getAppToken: () => Promise<string>;
  fetch: FetchLike;
}

export interface ColleagueWorkLocationsRequest {
  users: string[];
  startDateTime: string;
  endDateTime: string;
  timeZone?: string;
}

export interface SharedOccurrence {
  workLocationType: string;
  start: unknown;
  end: unknown;
  placeId?: string;
}

export type ColleagueResult =
  | {
      user: string;
      shared: true;
      detail: 'approximate' | 'specific';
      occurrences: SharedOccurrence[];
    }
  | { user: string; shared: false; reason: string }
  | { user: string; error: string };

export class CallerRejectedError extends Error {}

function decodeJwtClaims(token: string): Record<string, unknown> | undefined {
  const parts = token.split('.');
  if (parts.length !== 3) return undefined;
  try {
    return JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf-8'));
  } catch {
    return undefined;
  }
}

/**
 * Accepts the caller only with a token Graph accepts (GET /me) whose claims name this
 * tenant and this app. Graph rejects a token whose signature does not verify, so the
 * claims of a token it accepts can be trusted.
 */
export async function verifyCaller(
  callerToken: string | undefined,
  deps: Pick<ColleagueWorkLocationsDeps, 'secrets' | 'graphBaseUrl' | 'fetch'>
): Promise<void> {
  if (!callerToken) throw new CallerRejectedError('No signed-in user for this request.');
  const claims = decodeJwtClaims(callerToken);
  const tenant = deps.secrets.tenantId?.toLowerCase();
  const clientId = deps.secrets.clientId?.toLowerCase();
  const tokenTenant = String(claims?.tid ?? '').toLowerCase();
  const tokenApp = String(claims?.appid ?? claims?.azp ?? '').toLowerCase();
  if (!claims || !tenant || tokenTenant !== tenant || tokenApp !== clientId) {
    throw new CallerRejectedError('The access token was not issued to this app in this tenant.');
  }
  const response = await deps.fetch(`${deps.graphBaseUrl}/v1.0/me?$select=id`, {
    headers: { Authorization: `Bearer ${callerToken}` },
  });
  if (!response.ok) {
    throw new CallerRejectedError(`Graph rejected the caller's token (HTTP ${response.status}).`);
  }
}

export function validateRequest(request: ColleagueWorkLocationsRequest): string | undefined {
  const users = request.users ?? [];
  if (users.length === 0) return 'Name at least one user (email or user id).';
  if (users.length > MAX_USERS) return `Name at most ${MAX_USERS} users per call.`;
  if (users.some((u) => typeof u !== 'string' || !/^[^\s/?#'()]+$/.test(u))) {
    return 'Each user must be an email address (UPN) or a user id.';
  }
  const start = Date.parse(request.startDateTime);
  const end = Date.parse(request.endDateTime);
  if (Number.isNaN(start) || Number.isNaN(end)) {
    return 'startDateTime and endDateTime must be ISO 8601 date-times.';
  }
  if (end <= start) return 'endDateTime must be after startDateTime.';
  if (end - start > MAX_RANGE_DAYS * 24 * 3600 * 1000) {
    return `The range can span at most ${MAX_RANGE_DAYS} days.`;
  }
  if (request.timeZone !== undefined && !/^[\w +\-/().]{1,64}$/.test(request.timeZone)) {
    return 'timeZone must be a Windows or IANA time zone name.';
  }
  return undefined;
}

type GraphResult =
  | { ok: true; body: Record<string, unknown> }
  | { ok: false; status: number; message?: string };

async function graphGet(
  deps: ColleagueWorkLocationsDeps,
  token: string,
  path: string,
  timeZone?: string
): Promise<GraphResult> {
  const headers: Record<string, string> = { Authorization: `Bearer ${token}` };
  if (timeZone) headers.Prefer = `outlook.timezone="${timeZone}"`;
  const response = await deps.fetch(`${deps.graphBaseUrl}/v1.0${path}`, { headers });
  if (!response.ok) {
    let message: string | undefined;
    try {
      const body = (await response.json()) as { error?: { message?: unknown } };
      if (typeof body.error?.message === 'string') message = body.error.message.slice(0, 300);
    } catch {
      /* no JSON body */
    }
    return { ok: false, status: response.status, message };
  }
  return { ok: true, body: (await response.json()) as Record<string, unknown> };
}

function failure(result: { status: number; message?: string }): string {
  const detail = result.message ? ` ${result.message}` : '';
  if (result.status === 404) return `User not found.${detail}`;
  if (result.status === 401 || result.status === 403) {
    return `Not readable: the user is outside the mailboxes this server may read.${detail}`;
  }
  return `Graph returned HTTP ${result.status}.${detail}`;
}

const OBJECT_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * The user's object id. occurrencesView answers HTTP 500 ("Invalid smtp address") when
 * the user is named by UPN, so an email is resolved first, with the caller's own token
 * (delegated User.Read.All): the app token needs no directory permission.
 */
async function resolveObjectId(
  deps: ColleagueWorkLocationsDeps,
  callerToken: string,
  user: string
): Promise<{ id: string } | { error: string }> {
  if (OBJECT_ID.test(user)) return { id: user };
  const result = await graphGet(deps, callerToken, `/users/${encodeURIComponent(user)}?$select=id`);
  if (!result.ok) return { error: failure(result) };
  const id = result.body.id;
  return typeof id === 'string' && OBJECT_ID.test(id) ? { id } : { error: 'User not found.' };
}

async function readColleague(
  deps: ColleagueWorkLocationsDeps,
  callerToken: string,
  appToken: string,
  user: string,
  request: ColleagueWorkLocationsRequest
): Promise<ColleagueResult> {
  const resolved = await resolveObjectId(deps, callerToken, user);
  if ('error' in resolved) return { user, error: resolved.error };
  const base = `/users/${resolved.id}/settings/workHoursAndLocations`;

  // The colleague's own sharing level. Without it the server cannot honour `none`, so it
  // returns nothing for that user.
  const setting = await graphGet(deps, appToken, `${base}?$select=maxSharedWorkLocationDetails`);
  if (!setting.ok) return { user, error: failure(setting) };
  const level = String(setting.body.maxSharedWorkLocationDetails ?? '');
  if (level === 'none') {
    return { user, shared: false, reason: 'The user does not share their work location.' };
  }
  const detail = level === 'specific' ? 'specific' : 'approximate';

  const start = encodeURIComponent(request.startDateTime);
  const end = encodeURIComponent(request.endDateTime);
  const view = await graphGet(
    deps,
    appToken,
    `${base}/occurrencesView(startDateTime='${start}',endDateTime='${end}')` +
      '?$select=workLocationType,start,end,placeId',
    request.timeZone
  );
  if (!view.ok) return { user, error: failure(view) };

  const values = Array.isArray(view.body.value)
    ? (view.body.value as Record<string, unknown>[])
    : [];
  const occurrences = values.map((o): SharedOccurrence => {
    const shared: SharedOccurrence = {
      workLocationType: String(o.workLocationType ?? 'unspecified'),
      start: o.start,
      end: o.end,
    };
    if (detail === 'specific' && typeof o.placeId === 'string' && o.placeId) {
      shared.placeId = o.placeId;
    }
    return shared;
  });
  return { user, shared: true, detail, occurrences };
}

export async function getColleagueWorkLocations(
  callerToken: string | undefined,
  request: ColleagueWorkLocationsRequest,
  deps: ColleagueWorkLocationsDeps
): Promise<ColleagueResult[]> {
  await verifyCaller(callerToken, deps);
  const appToken = await deps.getAppToken();
  const users = [...new Set(request.users.map((u) => u.trim()))];
  return Promise.all(
    users.map((user) =>
      readColleague(deps, callerToken as string, appToken, user, request).catch(
        (error: Error): ColleagueResult => ({ user, error: error.message })
      )
    )
  );
}

let defaultDeps: Promise<ColleagueWorkLocationsDeps> | undefined;

/** Production dependencies: the server's secrets, an app token client, global fetch. */
export function getDefaultDeps(): Promise<ColleagueWorkLocationsDeps> {
  defaultDeps ??= getSecrets().then((secrets) => {
    const client = new AppTokenClient(secrets);
    return {
      secrets,
      graphBaseUrl: getGraphBaseUrl(secrets.cloudType),
      getAppToken: () => client.getToken(),
      fetch: (url, init) => fetch(url, init),
    };
  });
  defaultDeps.catch(() => {
    defaultDeps = undefined;
  });
  return defaultDeps;
}
