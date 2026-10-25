// `@`-mentions in a comment's plain text, format-neutral: the spans a comments pane draws and the
// typeahead that adds them. A mention is an offset span into the text; products map their own
// records onto it (PowerPoint's modern comment mentions in `pptx/render/comment-mentions`).
// DOM-free.

/** One mention: `length` characters of the text from `start`, the `@` included. */
export interface OfficeCommentMention {
	start: number;
	length: number;
	/** The mentioned person's id in the product's people list. */
	id?: string;
	/** The mentioned person's display name. */
	name?: string;
}

/** A person the typeahead offers. */
export interface OfficeMentionCandidate {
	id: string;
	name: string;
	initials?: string;
}

/** One run of a comment's text: plain text or a mention. */
export interface OfficeCommentSegment {
	kind: 'text' | 'mention';
	text: string;
	id?: string;
	name?: string;
}

const usable = (text: string, mentions: readonly OfficeCommentMention[] | undefined) =>
	(mentions ?? [])
		.filter(
			(m) =>
				Number.isFinite(m.start) &&
				Number.isFinite(m.length) &&
				m.length > 0 &&
				m.start >= 0 &&
				m.start + m.length <= text.length,
		)
		.sort((a, b) => a.start - b.start);

/**
 * The text split into plain runs and mentions. Spans out of range, empty or overlapping an earlier
 * one are ignored (the offsets can come from a file another program edited).
 */
export function commentSegments(
	text: string,
	mentions?: readonly OfficeCommentMention[],
): OfficeCommentSegment[] {
	const body = String(text ?? '');
	const segments: OfficeCommentSegment[] = [];
	let cursor = 0;
	for (const mention of usable(body, mentions)) {
		if (mention.start < cursor) continue;
		if (mention.start > cursor)
			segments.push({ kind: 'text', text: body.slice(cursor, mention.start) });
		segments.push({
			kind: 'mention',
			text: body.slice(mention.start, mention.start + mention.length),
			...(mention.id !== undefined ? { id: mention.id } : {}),
			...(mention.name !== undefined ? { name: mention.name } : {}),
		});
		cursor = mention.start + mention.length;
	}
	if (cursor < body.length) segments.push({ kind: 'text', text: body.slice(cursor) });
	return segments;
}

/** The `@` token being typed: the text after the `@`, lower-cased, and where the `@` is. */
export interface OfficeMentionQuery {
	query: string;
	start: number;
}

/**
 * The `@` token just left of `caret`, or `null`. The `@` must start the text or follow whitespace and
 * the token holds no whitespace, so an e-mail address does not open the typeahead.
 */
export function mentionQueryAt(text: string, caret: number): OfficeMentionQuery | null {
	const body = String(text ?? '');
	const end = Math.max(0, Math.min(caret, body.length));
	for (let index = end - 1; index >= 0; index -= 1) {
		const character = body[index]!;
		if (character === '@') {
			if (index > 0 && !/\s/u.test(body[index - 1]!)) return null;
			return { query: body.slice(index + 1, end).toLowerCase(), start: index };
		}
		if (/\s/u.test(character)) return null;
	}
	return null;
}

const score = (candidate: OfficeMentionCandidate, needle: string): number => {
	if (!needle) return 1;
	const name = candidate.name.toLowerCase();
	if (name.startsWith(needle)) return 0;
	if (name.includes(needle) || (candidate.initials ?? '').toLowerCase().startsWith(needle))
		return 2;
	return -1;
};

/** The candidates a query matches, best first: name prefixes, then names or initials containing it. */
export function matchMentionCandidates<C extends OfficeMentionCandidate>(
	candidates: readonly C[],
	query: string,
	limit = 6,
): C[] {
	const needle = String(query ?? '')
		.trim()
		.toLowerCase();
	return candidates
		.filter((candidate) => candidate.id && candidate.name)
		.map((candidate) => ({ candidate, rank: score(candidate, needle) }))
		.filter(({ rank }) => rank >= 0)
		.sort((a, b) => a.rank - b.rank || a.candidate.name.localeCompare(b.candidate.name))
		.slice(0, limit)
		.map(({ candidate }) => candidate);
}

/** The text, mentions and caret after a pick from the typeahead. */
export interface OfficeMentionInsertion {
	text: string;
	mentions: OfficeCommentMention[];
	caret: number;
}

/**
 * The `@` token at `caret` (or the caret itself) replaced by `@<name>` and a space, with the new span
 * (the `@` included, as Office records it) and the spans after it shifted. A span the token overlaps
 * (a pick made inside an existing mention) is dropped, so the text is never claimed twice.
 */
export function insertMention(
	text: string,
	mentions: readonly OfficeCommentMention[] | undefined,
	caret: number,
	candidate: OfficeMentionCandidate,
): OfficeMentionInsertion {
	const body = String(text ?? '');
	const at = Math.max(0, Math.min(caret, body.length));
	const active = mentionQueryAt(body, at);
	// The token runs from its @ to the caret; the lower-cased query can be longer (İ is i and a dot).
	const start = active ? active.start : at;
	const end = at;
	const label = `@${candidate.name}`;
	const delta = label.length + 1 - (end - start);
	const next = (mentions ?? [])
		.filter((m) => m.start >= end || m.start + m.length <= start)
		.map((m) => (m.start >= end ? { ...m, start: m.start + delta } : m));
	next.push({ start, length: label.length, id: candidate.id, name: candidate.name });
	next.sort((a, b) => a.start - b.start);
	return {
		text: `${body.slice(0, start)}${label} ${body.slice(end)}`,
		mentions: next,
		caret: start + label.length + 1,
	};
}

/** What a box posts: its text without the surrounding whitespace, and the spans moved with it. */
export function trimMentionText(
	text: string,
	mentions: readonly OfficeCommentMention[] | undefined,
): { text: string; mentions: OfficeCommentMention[] } {
	const body = String(text ?? '');
	const trimmed = body.trim();
	const lead = trimmed ? body.length - body.trimStart().length : 0;
	return {
		text: trimmed,
		mentions: usable(body, mentions)
			.filter((m) => m.start >= lead && m.start + m.length <= lead + trimmed.length)
			.map((m) => ({ ...m, start: m.start - lead })),
	};
}

/**
 * The spans after the text changed from `before` to `after` by one edit (typing, deleting or a
 * paste): spans before the changed part stay, spans after it move by the length change, and a span
 * the edit touches is dropped. The same array when the text did not change.
 */
export function followTextEdit(
	before: string,
	after: string,
	mentions: readonly OfficeCommentMention[],
): readonly OfficeCommentMention[] {
	if (before === after) return mentions;
	let prefix = 0;
	while (prefix < before.length && prefix < after.length && before[prefix] === after[prefix])
		prefix += 1;
	let suffix = 0;
	while (
		suffix < before.length - prefix &&
		suffix < after.length - prefix &&
		before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
	)
		suffix += 1;
	const changedEnd = before.length - suffix;
	const delta = after.length - before.length;
	return mentions.flatMap((m) => {
		if (m.start + m.length <= prefix) return [m];
		if (m.start >= changedEnd) return [{ ...m, start: m.start + delta }];
		return [];
	});
}
