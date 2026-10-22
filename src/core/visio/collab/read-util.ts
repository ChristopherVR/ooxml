// Small pure helpers of the room reader: the change-detecting hash, key splitting and connects.
import { compare } from './read-ids';
import type { ResolvedShape } from './read-shapes';
import type { ConnectEntry } from './schema';

export interface ResolvedConnect {
	key: string;
	entry: ConnectEntry;
	/** Final ids the connect refers to. */
	from: string;
	to: string;
}
/** A 64-bit FNV-style hash of the strings, as hex: a change detector, not a security measure. */
export function hashTexts(texts: Iterable<string>): string {
	let a = 0x811c9dc5;
	let b = 0x01000193 ^ 0x9e3779b9;
	for (const text of texts) {
		for (let i = 0; i < text.length; i++) {
			const c = text.charCodeAt(i);
			a = Math.imul(a ^ c, 0x01000193);
			b = Math.imul(b + c, 0x85ebca6b) ^ (b >>> 13);
		}
		a = Math.imul(a ^ 0xff, 0x01000193);
		b = Math.imul(b ^ 0xff, 0xc2b2ae35);
	}
	return `${(a >>> 0).toString(16)}${(b >>> 0).toString(16)}`;
}

/** Splits `key` at its first `depth` slashes; the last part keeps any further slashes. */
export const split = (key: string, depth: number): string[] | undefined => {
	const out: string[] = [];
	let from = 0;
	for (let i = 0; i < depth; i++) {
		const slash = key.indexOf('/', from);
		if (slash < 0) return undefined;
		out.push(key.slice(from, slash));
		from = slash + 1;
	}
	out.push(key.slice(from));
	return out;
};

export const connectSignature = (connects: readonly ResolvedConnect[]): string[] =>
	connects.flatMap((c) => [c.from, c.to, c.entry.fc, c.entry.tc, c.entry.fp, c.entry.tp]);

/** The connects of one page or master whose ends are alive, with their final shape ids. */
export function resolveConnects(
	list: { key: string; entry: ConnectEntry }[] | undefined,
	shapes: readonly ResolvedShape[],
): ResolvedConnect[] {
	const byKey = new Map(shapes.map((s) => [s.key, s]));
	const out: ResolvedConnect[] = [];
	for (const item of [...(list ?? [])].sort((a, b) => compare(a.key, b.key))) {
		const { entry } = item;
		const from = entry.fk ? byKey.get(entry.fk) : undefined;
		const to = entry.tk ? byKey.get(entry.tk) : undefined;
		if ((entry.fk && !from) || (entry.tk && !to)) continue;
		out.push({ key: item.key, entry, from: from?.id ?? entry.fi, to: to?.id ?? entry.ti });
	}
	return out;
}
