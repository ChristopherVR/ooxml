// Cell formats in a shared workbook. Style ids are indexes into one workbook's `styles` array, so
// they mean nothing to another peer (two peers creating a format at once would both claim the
// next index). The shared table is content-addressed instead: a format's key is a hash of its
// canonical JSON (`styleKey`), entries are immutable, and concurrent creation of the same format
// writes the same key. Each peer interns the formats it reads into its own `styles` array.
import type * as Y from 'yjs';
import type { CellStyle, Workbook } from '../model';
import { internStyle, styleKey } from '../styles';
import { isRecord } from './schema';

/** cyrb53: a fast 53-bit string hash, ample for the few hundred formats a workbook holds. */
function hash53(text: string): number {
	let h1 = 0xdeadbeef;
	let h2 = 0x41c6ce57;
	for (let i = 0; i < text.length; i++) {
		const ch = text.charCodeAt(i);
		h1 = Math.imul(h1 ^ ch, 2654435761);
		h2 = Math.imul(h2 ^ ch, 1597334677);
	}
	h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
	h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
	return 4294967296 * (2097151 & h2) + (h1 >>> 0);
}

const hashes = new WeakMap<CellStyle, string>();

/** The shared key of a format. */
export function sharedStyleKey(style: CellStyle): string {
	let key = hashes.get(style);
	if (key === undefined) {
		key = hash53(styleKey(style)).toString(36);
		hashes.set(style, key);
	}
	return key;
}

/** Light structural check of a format read from a peer. */
export function isCellStyle(value: unknown): value is CellStyle {
	return (
		isRecord(value) &&
		isRecord(value.font) &&
		isRecord(value.fill) &&
		isRecord(value.border) &&
		typeof value.numFmt === 'string'
	);
}

/**
 * Translates between local style ids and shared keys. A cell whose format equals the room's
 * default format carries no key.
 */
export class StyleBridge {
	private readonly ids = new Map<string, number | undefined>();

	constructor(
		private readonly workbook: Workbook,
		private readonly table: Y.Map<unknown>,
		readonly defaultKey: string,
		/** Writers publish the formats they reference; readers never write. */
		private readonly writable: boolean,
	) {}

	/** The shared key of a local style id (undefined for the room default). */
	keyOf = (styleId: number | undefined): string | undefined => {
		const style = this.workbook.styles[styleId ?? 0] ?? this.workbook.styles[0];
		if (!style) return undefined;
		const key = sharedStyleKey(style);
		if (this.writable && !this.table.has(key)) this.table.set(key, structuredClone(style));
		return key === this.defaultKey ? undefined : key;
	};

	/** The local style id of a shared key (undefined for style 0). */
	idOf = (key: string | undefined): number | undefined => {
		const resolved = key ?? this.defaultKey;
		if (this.ids.has(resolved)) return this.ids.get(resolved);
		const style = this.table.get(resolved);
		const id = isCellStyle(style) ? internStyle(this.workbook, structuredClone(style)) : 0;
		const out = id === 0 ? undefined : id;
		this.ids.set(resolved, out);
		return out;
	};
}
