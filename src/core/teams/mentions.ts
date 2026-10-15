// @mentions. A mention is stored next to the plain message text as a structured range
// (`{ kind, id, name, start, end }`), so a client that has never heard of mentions still shows the
// message as ordinary text ("hi @Ada Lovelace"). Offsets are always derived from the final cleaned
// text and re-validated when read back, so a peer cannot point a mention at text that is not there.
// Pure and DOM-free: the picker, the highlighting and the activity feed all share these rules.
import { isValidId, stripHtmlTags } from '../collab/validation';
import type { Message } from './model';

/** `user` points at one person; `channel` and `team` address everyone (`@channel`, `@team`). */
export type MentionKind = 'user' | 'channel' | 'team';

/** Someone (or everyone) that can be mentioned, as offered by the picker. */
export interface MentionTarget {
	kind: MentionKind;
	/** The user id; `channel` / `team` for the group mentions. */
	id: string;
	/** The display name that follows the `@`. */
	name: string;
}

/** A mention located in a message: `text.slice(start, end)` is `@` plus the name. */
export interface Mention extends MentionTarget {
	start: number;
	end: number;
}

export interface MentionSegment {
	text: string;
	mention?: Mention;
}

export const MAX_MENTIONS = 50;
const MAX_NAME_CHARS = 80;
const MAX_QUERY_CHARS = 40;

/** `@channel` and `@team`, offered after the people. */
export const GROUP_MENTIONS: readonly MentionTarget[] = [
	{ kind: 'channel', id: 'channel', name: 'channel' },
	{ kind: 'team', id: 'team', name: 'team' },
];

const WORD = /[\p{L}\p{N}_]/u;
const isWord = (ch: string | undefined): boolean => ch !== undefined && ch !== '' && WORD.test(ch);

/** A mention name: markup removed, whitespace collapsed, no `@`, bounded. Empty when unusable. */
export function sanitizeMentionName(name: unknown): string {
	if (typeof name !== 'string') return '';
	return stripHtmlTags(name)
		.replace(/[\u0000-\u001f\u007f@]/gu, ' ')
		.replace(/\s+/gu, ' ')
		.trim()
		.slice(0, MAX_NAME_CHARS);
}

/** Validate a target from an untrusted source (a peer's row, a host's property). */
export function sanitizeMentionTarget(raw: unknown): MentionTarget | null {
	if (!raw || typeof raw !== 'object') return null;
	const r = raw as Record<string, unknown>;
	const name = sanitizeMentionName(r.name);
	if (!name) return null;
	if (r.kind === 'channel' || r.kind === 'team')
		return name === r.kind ? { kind: r.kind, id: r.kind, name } : null;
	if (r.kind !== 'user' || !isValidId(r.id)) return null;
	return { kind: 'user', id: r.id, name };
}

/**
 * Find the mentions in `text`. Candidates are tried in the given order for ties; the longest name
 * wins at a position (`@Ada Lovelace` beats `@Ada`). A mention needs a non-word character (or the
 * edge of the text) on both sides, so an e-mail address is not a mention. Case-insensitive.
 */
export function resolveMentions(text: string, candidates: readonly MentionTarget[]): Mention[] {
	const targets: MentionTarget[] = [];
	for (const candidate of candidates) {
		const clean = sanitizeMentionTarget(candidate);
		if (clean) targets.push(clean);
	}
	if (targets.length === 0) return [];
	const lower = text.toLowerCase();
	const out: Mention[] = [];
	let from = 0;
	for (let at = text.indexOf('@', from); at >= 0 && out.length < MAX_MENTIONS;) {
		let best: MentionTarget | undefined;
		if (!isWord(text[at - 1]))
			for (const target of targets) {
				const end = at + 1 + target.name.length;
				if (end > text.length || isWord(text[end])) continue;
				if (lower.slice(at + 1, end) !== target.name.toLowerCase()) continue;
				if (!best || target.name.length > best.name.length) best = target;
			}
		if (best) {
			const end = at + 1 + best.name.length;
			out.push({ ...best, start: at, end });
			from = end;
		} else from = at + 1;
		at = text.indexOf('@', from);
	}
	return out;
}

/** Re-validate stored mentions against the text they claim to sit in. */
export function sanitizeMentions(raw: unknown, text: string): Mention[] {
	if (!Array.isArray(raw)) return [];
	const out: Mention[] = [];
	let floor = 0;
	for (const item of raw.slice(0, MAX_MENTIONS)) {
		const target = sanitizeMentionTarget(item);
		const { start, end } = (item ?? {}) as { start?: unknown; end?: unknown };
		if (!target || typeof start !== 'number' || typeof end !== 'number') continue;
		if (!Number.isInteger(start) || !Number.isInteger(end) || start < floor || end > text.length)
			continue;
		if (text.slice(start, end).toLowerCase() !== `@${target.name}`.toLowerCase()) continue;
		out.push({ ...target, start, end });
		floor = end;
	}
	return out;
}

/** Split text into plain runs and mention runs, for rendering. */
export function mentionSegments(text: string, mentions: readonly Mention[] = []): MentionSegment[] {
	const out: MentionSegment[] = [];
	let at = 0;
	for (const mention of mentions) {
		if (mention.start < at || mention.end > text.length) continue;
		if (mention.start > at) out.push({ text: text.slice(at, mention.start) });
		out.push({ text: text.slice(mention.start, mention.end), mention });
		at = mention.end;
	}
	if (at < text.length) out.push({ text: text.slice(at) });
	return out;
}

/** Whether a message addresses this person: by name, or through `@channel` / `@team`. */
export function mentionsUser(message: Pick<Message, 'mentions'>, userId: string): boolean {
	return (message.mentions ?? []).some(
		(m) => m.kind === 'channel' || m.kind === 'team' || m.id === userId,
	);
}

/** An open `@query` at the caret, as the composer's picker needs it. */
export interface MentionQuery {
	/** Index of the `@`. */
	start: number;
	/** What was typed after it. */
	query: string;
}

/** The `@query` being typed right before `caret`, or null. The `@` must start a word. */
export function mentionQueryAt(text: string, caret: number): MentionQuery | null {
	const head = text.slice(0, Math.max(0, Math.min(caret, text.length)));
	const start = head.lastIndexOf('@');
	if (start < 0 || isWord(head[start - 1])) return null;
	const query = head.slice(start + 1);
	if (query.length > MAX_QUERY_CHARS || /[\n\r@]/u.test(query) || /\s{2}|\s$/u.test(query))
		return null;
	return { start, query };
}

/** Targets whose name matches what was typed; names starting with it (or a word of it) first. */
export function filterMentionTargets(
	targets: readonly MentionTarget[],
	query: string,
	limit = 8,
): MentionTarget[] {
	const q = query.trim().toLowerCase();
	const rank = (target: MentionTarget): number => {
		const name = target.name.toLowerCase();
		if (!q || name.startsWith(q)) return 0;
		if (name.split(' ').some((word) => word.startsWith(q))) return 1;
		return name.includes(q) ? 2 : 3;
	};
	return targets
		.map((target, index) => ({ target, index, r: rank(target) }))
		.filter((entry) => entry.r < 3)
		.sort((a, b) => a.r - b.r || a.index - b.index)
		.slice(0, limit)
		.map((entry) => entry.target);
}

/** Replace the open query with the mention token (and a trailing space) and place the caret. */
export function insertMention(
	text: string,
	query: MentionQuery,
	caret: number,
	target: MentionTarget,
): { text: string; caret: number } {
	const token = `@${target.name} `;
	const rest = text.slice(caret);
	const tail = rest.startsWith(' ') ? rest.slice(1) : rest;
	return { text: text.slice(0, query.start) + token + tail, caret: query.start + token.length };
}

/**
 * Who the picker offers in a channel: people online now, people who wrote there, then the group
 * mentions. The author of the draft is left out.
 */
export function mentionTargetsFor(
	messages: readonly Pick<Message, 'authorId' | 'authorName'>[],
	people: readonly { id: string; name: string }[],
	selfId: string,
): MentionTarget[] {
	const byId = new Map<string, MentionTarget>();
	const add = (id: string, rawName: string): void => {
		const target = sanitizeMentionTarget({ kind: 'user', id, name: rawName });
		if (target && id !== selfId && !byId.has(id)) byId.set(id, target);
	};
	for (const person of people) add(person.id, person.name);
	for (const message of messages) add(message.authorId, message.authorName);
	const users = [...byId.values()].sort((a, b) => a.name.localeCompare(b.name));
	return [...users, ...GROUP_MENTIONS];
}
