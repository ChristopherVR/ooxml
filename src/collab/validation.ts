// Validation and sanitisation of everything that arrives from a peer or a host: room ids, display
// names, colours, avatar URLs and numeric coordinates. Pure functions, no DOM.
//
// Extracted from pptx-viewer `collaboration-presence.ts` (room id, name, colour, avatar, mixed
// content) and docx-viewer `collaboration-protocol.ts` (`validId`); see PROVENANCE.md.

/** Fallback cursor/label colour when none is supplied or it fails validation. */
export const DEFAULT_COLOR = '#4c8bf5';

const ROOM_ID_REGEX = /^[a-zA-Z0-9_-]{1,128}$/u;
const HEX_COLOR_REGEX = /^#[0-9a-fA-F]{6}$/u;
// Control characters are never printable in a display name.
// eslint-disable-next-line no-control-regex
const CONTROL_CHARS = /[\u0000-\u001f\u007f]/u;

/** True when `roomId` is a safe 1-128 char alphanumeric/`-`/`_` token. */
export function isValidRoomId(roomId: unknown): roomId is string {
	return typeof roomId === 'string' && ROOM_ID_REGEX.test(roomId);
}

/** Return `roomId` when valid, throw otherwise. */
export function validateRoomId(roomId: string): string {
	if (!isValidRoomId(roomId)) {
		throw new Error(
			`Invalid collaboration room ID: "${String(roomId)}". Must be 1-128 alphanumeric characters, hyphens, or underscores.`,
		);
	}
	return roomId;
}

/** True for a non-empty string of at most `max` characters (session, client and batch ids). */
export function isValidId(value: unknown, max = 160): value is string {
	return typeof value === 'string' && value.length > 0 && value.length <= max;
}

/** True for a non-negative safe integer (versions, sequence numbers, offsets). */
export function isNonNegativeSafeInteger(value: unknown): value is number {
	return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

/**
 * Remove HTML tags in a single linear pass with no regex backtracking. Tracking open/closed tag
 * state while walking the string means nested or overlapping `<`/`>` sequences (`<scr<script>ipt>`)
 * can never leave a live tag behind, and there is no polynomial ReDoS on many unmatched `<`.
 */
export function stripHtmlTags(text: string): string {
	let result = '';
	let insideTag = false;
	for (const char of text) {
		if (char === '<') insideTag = true;
		else if (char === '>') insideTag = false;
		else if (!insideTag) result += char;
	}
	return result;
}

/** Strip HTML tags, trim, and clamp to 64 characters; falls back to `fallback`. */
export function sanitizeUserName(name: unknown, fallback = 'Anonymous'): string {
	if (typeof name !== 'string') return fallback;
	const trimmed = stripHtmlTags(name).trim().slice(0, 64);
	return trimmed || fallback;
}

/**
 * Strict profile-name check for hosts that prefer rejection over repair: 1 to 80 printable
 * characters. Returns the trimmed name or throws.
 */
export function validateDisplayName(name: string): string {
	const trimmed = name.trim();
	if (!trimmed || trimmed.length > 80 || CONTROL_CHARS.test(trimmed))
		throw new Error('Display name must contain 1 to 80 printable characters');
	return trimmed;
}

/** Validate a 6-digit hex colour; returns `fallback` when invalid. */
export function sanitizeColor(color: unknown, fallback: string = DEFAULT_COLOR): string {
	return typeof color === 'string' && HEX_COLOR_REGEX.test(color) ? color : fallback;
}

/** Allow only http(s)/data: avatar URLs; otherwise `undefined`. */
export function sanitizeAvatarUrl(url: unknown): string | undefined {
	if (typeof url !== 'string') return undefined;
	try {
		const { protocol } = new URL(url);
		if (protocol === 'https:' || protocol === 'http:' || protocol === 'data:') return url;
	} catch {
		// not an absolute URL: reject
	}
	return undefined;
}

/** Coerce to a non-negative integer index (slide, page, sheet). */
export function sanitizeIndex(value: unknown): number {
	if (typeof value !== 'number' || !Number.isFinite(value)) return 0;
	return Math.max(0, Math.floor(value));
}

/** Clamp a coordinate to `[min - margin, max + margin]`; non-finite input becomes 0. */
export function clampCoordinate(value: unknown, min: number, max: number, margin = 20): number {
	if (typeof value !== 'number' || !Number.isFinite(value)) return 0;
	return Math.max(min - margin, Math.min(max + margin, value));
}

const LOOPBACK_HOSTS = ['localhost', '127.0.0.1', '[::1]'];

/**
 * True when connecting to `serverUrl` would be blocked as mixed content: an insecure `ws://` socket
 * opened from a secure `https:` page. Loopback hosts are exempt. `pageProtocol` defaults to the
 * ambient `location.protocol` when there is one (no DOM types are required).
 */
export function isMixedContentBlocked(
	serverUrl: string,
	pageProtocol: string | undefined = (globalThis as { location?: { protocol?: string } }).location
		?.protocol,
): boolean {
	if (pageProtocol !== 'https:') return false;
	let parsed: URL;
	try {
		parsed = new URL(serverUrl);
	} catch {
		return false;
	}
	return parsed.protocol === 'ws:' && !LOOPBACK_HOSTS.includes(parsed.hostname);
}
