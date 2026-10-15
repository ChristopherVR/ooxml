// Conversions between what the composer shows (a mention is just the person's name, highlighted,
// with the `@` only while the picker is open, as in Teams) and what is stored (`@Name` plus a
// structured range, so older clients still read the message). Pure and DOM-free.
import { type MentionTarget, sanitizeMentionTarget } from './mentions';
import type { Message } from './model';

const WORD = /[\p{L}\p{N}_]/u;
const isWord = (ch: string | undefined): boolean => ch !== undefined && ch !== '' && WORD.test(ch);

/**
 * Turn the names picked in the composer back into stored `@Name` text. Each pick claims one
 * occurrence (whole words, case-insensitive, longest name first, never one already preceded by
 * `@`), so a plain word that happens to match a name elsewhere in the message stays plain.
 */
export function promoteMentions(text: string, picks: readonly MentionTarget[]): string {
	const wanted = new Map<string, { target: MentionTarget; count: number }>();
	for (const pick of picks) {
		const target = sanitizeMentionTarget(pick);
		if (!target) continue;
		const key = `${target.kind}:${target.id}`;
		const entry = wanted.get(key);
		if (entry) entry.count++;
		else wanted.set(key, { target, count: 1 });
	}
	const lower = text.toLowerCase();
	const claimed: [number, number][] = [];
	const inserts: number[] = [];
	const free = (start: number, end: number): boolean =>
		!claimed.some(([a, b]) => start < b && end > a);
	const ordered = [...wanted.values()].sort((a, b) => b.target.name.length - a.target.name.length);
	for (const { target, count } of ordered) {
		const name = target.name.toLowerCase();
		let left = count;
		for (let at = lower.indexOf(name); at >= 0 && left > 0; at = lower.indexOf(name, at + 1)) {
			const end = at + name.length;
			if (isWord(text[at - 1]) || text[at - 1] === '@' || isWord(text[end]) || !free(at, end))
				continue;
			claimed.push([at, end]);
			inserts.push(at);
			left--;
		}
	}
	let out = text;
	for (const at of inserts.sort((a, b) => b - a)) out = `${out.slice(0, at)}@${out.slice(at)}`;
	return out;
}

/** The composer's view of a stored message: mention names without the `@`. */
export function mentionDisplayText(message: Pick<Message, 'text' | 'mentions'>): string {
	let out = '';
	let at = 0;
	for (const mention of message.mentions ?? []) {
		out +=
			message.text.slice(at, mention.start) + message.text.slice(mention.start + 1, mention.end);
		at = mention.end;
	}
	return out + message.text.slice(at);
}
