import { NS, children, first, parseXml } from '../../xml/index';
import { formatAddress } from '../address';
import type { CommentMessage, Workbook } from '../model';

/** A deterministic GUID-shaped id (`{XXXXXXXX-XXXX-...}`) from a seed string. */
export function guidFrom(seed: string): string {
	let h1 = 0x811c9dc5;
	let h2 = 0x01000193;
	const hex: string[] = [];
	for (let round = 0; round < 4; round++) {
		for (let i = 0; i < seed.length; i++) {
			h1 = Math.imul(h1 ^ seed.charCodeAt(i), 0x01000193) >>> 0;
			h2 = Math.imul(h2 ^ (seed.charCodeAt(i) + round), 0x5bd1e995) >>> 0;
		}
		h1 = Math.imul(h1 ^ round, 0x27d4eb2d) >>> 0;
		hex.push(((h1 ^ h2) >>> 0).toString(16).toUpperCase().padStart(8, '0'));
	}
	const s = hex.join('');
	return `{${s.slice(0, 8)}-${s.slice(8, 12)}-4${s.slice(13, 16)}-8${s.slice(17, 20)}-${s.slice(20, 32)}}`;
}

const validGuid = (id: string): boolean =>
	/^\{[\da-f]{8}(?:-[\da-f]{4}){3}-[\da-f]{12}\}$/i.test(id);

/** Workbook-wide identities: imported messages retain theirs and copies get distinct ids. */
export class CommentIds {
	private readonly reserved = new Map<string, object>();
	private readonly used = new Set<string>();
	private readonly mentions = new WeakMap<CommentMessage, object[]>();

	constructor(workbook?: Workbook) {
		// Prefer the source cell over a copied message that still carries its source fragment.
		for (const matching of [true, false])
			for (const sheet of workbook?.sheets ?? []) {
				for (const comment of sheet.comments)
					for (const message of [comment, ...(comment.replies ?? [])]) {
						if (comment.regenerateThreadIds) continue;
						if (!message.threadedXml) continue;
						const node = parseXml(message.threadedXml).documentElement;
						if ((node.getAttribute('ref') === formatAddress(comment.address)) !== matching)
							continue;
						this.reserve(node.getAttribute('id') ?? '', message);
						const mentions = first(node, 'mentions', NS.tc);
						if (mentions)
							children(mentions, 'mention', NS.tc).forEach((mention, index) => {
								this.reserve(
									mention.getAttribute('mentionId') ?? '',
									this.mentionOwner(message, index),
								);
							});
					}
			}
	}

	private reserve(id: string, owner: object): void {
		const key = id.toUpperCase();
		if (validGuid(id) && !this.reserved.has(key)) this.reserved.set(key, owner);
	}

	mentionOwner(message: CommentMessage, index: number): object {
		let owners = this.mentions.get(message);
		if (!owners) this.mentions.set(message, (owners = []));
		return (owners[index] ??= {});
	}

	claim(source: string, seed: string, owner: object, preserve = true): string {
		const key = source.toUpperCase();
		if (
			preserve &&
			validGuid(source) &&
			!this.used.has(key) &&
			(!this.reserved.has(key) || this.reserved.get(key) === owner)
		) {
			this.used.add(key);
			return source;
		}
		let index = 0;
		let id = guidFrom(seed);
		while (this.used.has(id) || this.reserved.has(id)) id = guidFrom(`${seed}:${++index}`);
		this.used.add(id);
		return id;
	}
}
