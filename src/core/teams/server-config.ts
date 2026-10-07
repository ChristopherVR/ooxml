// Bring-your-own-server configuration: where chat state syncs, where calls are signalled and which
// STUN/TURN servers carry the media. Everything is optional: with no server at all the workspace
// runs "local" (tabs of one browser over BroadcastChannel). Input is untrusted (a settings form, a
// URL fragment, a JSON file), so it is validated and every problem is reported, never thrown.
// New code.
import type { IceServer } from './peer';

export type ServerMode = 'local' | 'server';

export interface TeamsServerConfig {
	mode: ServerMode;
	/** Base URL of a y-websocket compatible sync server: `wss://host/sync` (the room is appended). */
	syncUrl?: string;
	/** Base URL of the signaling relay: `wss://host/signal` (the call room is appended). */
	signalingUrl?: string;
	iceServers: IceServer[];
	iceTransportPolicy?: 'all' | 'relay';
	/** Bearer token sent as `?token=` on both sockets. Treat it as a secret. */
	token?: string;
}

export interface ServerConfigResult {
	config: TeamsServerConfig;
	/** Human-readable problems; the offending field is left out of `config`. */
	issues: string[];
}

/** Public STUN only: enough on open networks, never enough behind symmetric NAT (add TURN). */
export const DEFAULT_ICE_SERVERS: readonly IceServer[] = [{ urls: 'stun:stun.l.google.com:19302' }];

const WS_URL = /^wss?:\/\/[^\s/?#]+(?::\d+)?(?:\/[^\s?#]*)?$/iu;
const ICE_URL = /^(stuns?|turns?):[^\s]+$/iu;

function wsUrl(value: unknown, field: string, issues: string[]): string | undefined {
	if (value === undefined || value === null || value === '') return undefined;
	if (typeof value !== 'string' || !WS_URL.test(value.trim())) {
		issues.push(`${field} must be a ws:// or wss:// URL without query or fragment`);
		return undefined;
	}
	return trimTrailingSlashes(value.trim());
}

function iceServer(raw: unknown, index: number, issues: string[]): IceServer | null {
	if (!raw || typeof raw !== 'object') {
		issues.push(`iceServers[${index}] must be an object`);
		return null;
	}
	const r = raw as Record<string, unknown>;
	const urls = (Array.isArray(r.urls) ? r.urls : [r.urls]).filter(
		(u): u is string => typeof u === 'string' && ICE_URL.test(u.trim()),
	);
	if (urls.length === 0) {
		issues.push(`iceServers[${index}].urls needs a stun: or turn: URL`);
		return null;
	}
	const turn = urls.some((u) => /^turns?:/iu.test(u));
	if (turn && (typeof r.username !== 'string' || typeof r.credential !== 'string')) {
		issues.push(`iceServers[${index}] is a TURN server and needs username and credential`);
		return null;
	}
	return {
		urls: urls.length === 1 ? (urls[0] as string) : urls,
		...(typeof r.username === 'string' ? { username: r.username } : {}),
		...(typeof r.credential === 'string' ? { credential: r.credential } : {}),
	};
}

/** Validate untrusted server settings into a usable config plus a list of issues. */
export function parseServerConfig(input: unknown): ServerConfigResult {
	const issues: string[] = [];
	const raw = (input && typeof input === 'object' ? input : {}) as Record<string, unknown>;
	const syncUrl = wsUrl(raw.syncUrl, 'syncUrl', issues);
	const signalingUrl = wsUrl(raw.signalingUrl, 'signalingUrl', issues);
	const iceRaw = Array.isArray(raw.iceServers) ? raw.iceServers : undefined;
	const ice = iceRaw
		? iceRaw.map((s, i) => iceServer(s, i, issues)).filter((s): s is IceServer => s !== null)
		: [];
	const mode: ServerMode = raw.mode === 'server' && (syncUrl || signalingUrl) ? 'server' : 'local';
	if (raw.mode === 'server' && mode === 'local')
		issues.push('mode "server" needs a syncUrl or a signalingUrl; falling back to local');
	const config: TeamsServerConfig = {
		mode,
		iceServers: iceRaw ? ice : [...DEFAULT_ICE_SERVERS],
	};
	if (syncUrl) config.syncUrl = syncUrl;
	if (signalingUrl) config.signalingUrl = signalingUrl;
	if (raw.iceTransportPolicy === 'relay' || raw.iceTransportPolicy === 'all')
		config.iceTransportPolicy = raw.iceTransportPolicy;
	if (typeof raw.token === 'string' && raw.token.length > 0 && raw.token.length <= 4096)
		config.token = raw.token;
	if (
		config.iceTransportPolicy === 'relay' &&
		!ice.some((s) => [s.urls].flat().some((u) => /^turns?:/iu.test(u)))
	)
		issues.push('iceTransportPolicy "relay" needs a TURN server, otherwise no call can connect');
	return { config, issues };
}

/** `base/room?token=...` for a sync or signaling endpoint. The room is URL-encoded. */
export function endpointUrl(base: string, roomId: string, token?: string): string {
	const url = `${trimTrailingSlashes(base)}/${encodeURIComponent(roomId)}`;
	return token ? `${url}?token=${encodeURIComponent(token)}` : url;
}

/** The local-only config: no server, same-browser tabs. */
export function localServerConfig(): TeamsServerConfig {
	return { mode: 'local', iceServers: [...DEFAULT_ICE_SERVERS] };
}

function trimTrailingSlashes(value: string): string {
	let end = value.length;
	while (end > 0 && value.charCodeAt(end - 1) === 47) end--;
	return value.slice(0, end);
}
