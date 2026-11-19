// A bounded structure scan of a Yjs (v1) update, run before Yjs decodes it.
//
// Yjs allocates one object per struct, per value of a JSON or `any` content and per delete range
// while it decodes an update, and a struct can be two bytes long. A byte limit alone therefore lets
// a frame of a few megabytes allocate millions of objects. This walks the encoding with nothing but
// a cursor, counts what Yjs would allocate and stops at the first limit crossed, so an update over
// budget is refused before any of it is decoded. The layout read here is the one Yjs writes with
// its v1 update encoder (`writeStructs`, `writeDeleteSet`) and lib0's `writeAny`. New code.

/** What one update may make Yjs allocate. See `docs/collab-area.md` for how these were chosen. */
export interface UpdateLimits {
	/** Structs plus the values inside JSON and `any` contents. */
	maxStructs?: number;
	/** Ranges of the delete set. */
	maxDeleteRanges?: number;
}

/**
 * Default allocation budget of one update. The largest legitimate updates measured (the first sync
 * of a large deck, document, workbook or chat room) stay far below it; see `docs/collab-area.md`.
 */
export const DEFAULT_UPDATE_LIMITS: Readonly<Required<UpdateLimits>> = {
	maxStructs: 2_000_000,
	maxDeleteRanges: 1_000_000,
};

export type UpdateScan =
	| { ok: true; structs: number; deleteRanges: number; bytesRead: number }
	/** `overLimit` is true when the update is well formed so far but over budget. */
	| { ok: false; reason: string; overLimit: boolean };

/** Nesting depth of an `any` value that lib0 can read without exhausting the stack. */
const MAX_ANY_DEPTH = 64;

class Overrun extends Error {
	constructor(
		message: string,
		readonly overLimit = false,
	) {
		super(message);
	}
}

class Cursor {
	pos = 0;
	constructor(readonly bytes: Uint8Array) {}

	byte(): number {
		if (this.pos >= this.bytes.length) throw new Overrun('unexpected end of update');
		return this.bytes[this.pos++]!;
	}

	/** An unsigned varint (lib0 `readVarUint`), refused beyond the safe integer range. */
	uint(): number {
		let value = 0;
		let multiplier = 1;
		for (;;) {
			const byte = this.byte();
			value += (byte & 0x7f) * multiplier;
			if (byte < 0x80) return value;
			multiplier *= 128;
			if (multiplier > Number.MAX_SAFE_INTEGER) throw new Overrun('varint out of range');
		}
	}

	/** A signed varint (lib0 `readVarInt`): the same continuation rule, so skip it the same way. */
	skipInt(): void {
		let length = 0;
		while (this.byte() >= 0x80) {
			if (++length > 9) throw new Overrun('varint out of range');
		}
	}

	skip(length: number): void {
		if (length > this.bytes.length - this.pos) throw new Overrun('unexpected end of update');
		this.pos += length;
	}

	/** A length-prefixed string or byte array. */
	skipBuffer(): void {
		this.skip(this.uint());
	}
}

/** Tracks the allocation budget; throws once it is spent. */
class Budget {
	structs = 0;
	deleteRanges = 0;
	constructor(private readonly limits: Required<UpdateLimits>) {}

	addStructs(count: number): void {
		this.structs += count;
		if (this.structs > this.limits.maxStructs)
			throw new Overrun(`update holds more than ${this.limits.maxStructs} structs`, true);
	}

	addDeleteRanges(count: number): void {
		this.deleteRanges += count;
		if (this.deleteRanges > this.limits.maxDeleteRanges)
			throw new Overrun(
				`update holds more than ${this.limits.maxDeleteRanges} deleted ranges`,
				true,
			);
	}
}

/** One lib0 `writeAny` value. */
function skipAny(cursor: Cursor, budget: Budget, depth: number): void {
	if (depth > MAX_ANY_DEPTH) throw new Overrun('value nested too deeply');
	const tag = cursor.byte();
	switch (tag) {
		case 127: // undefined
		case 126: // null
		case 121: // false
		case 120: // true
			return;
		case 125: // integer
			return cursor.skipInt();
		case 124: // float32
			return cursor.skip(4);
		case 123: // float64
		case 122: // bigint64
			return cursor.skip(8);
		case 119: // string
		case 116: // Uint8Array
			return cursor.skipBuffer();
		case 118: {
			// object: key, value pairs
			const length = cursor.uint();
			budget.addStructs(length);
			for (let i = 0; i < length; i++) {
				cursor.skipBuffer();
				skipAny(cursor, budget, depth + 1);
			}
			return;
		}
		case 117: {
			const length = cursor.uint();
			budget.addStructs(length);
			for (let i = 0; i < length; i++) skipAny(cursor, budget, depth + 1);
			return;
		}
		default:
			throw new Overrun(`unknown value type ${tag}`);
	}
}

const BIT6 = 0x20;
const BIT7 = 0x40;
const BIT8 = 0x80;
const BITS5 = 0x1f;
/** Shared types whose content carries a name: `YXmlElement` and `YXmlHook`. */
const NAMED_TYPE_REFS = new Set([3, 5]);

/** The content of an item, by its content ref (Yjs `contentRefs`). */
function skipContent(cursor: Cursor, budget: Budget, ref: number): void {
	switch (ref) {
		case 1: // deleted
			cursor.uint();
			return;
		case 2: {
			// JSON: one string per value
			const length = cursor.uint();
			budget.addStructs(length);
			for (let i = 0; i < length; i++) cursor.skipBuffer();
			return;
		}
		case 3: // binary
		case 4: // string
		case 5: // embed, JSON as a string
			cursor.skipBuffer();
			return;
		case 6: // format: key and JSON value
			cursor.skipBuffer();
			cursor.skipBuffer();
			return;
		case 7: // type
			if (NAMED_TYPE_REFS.has(cursor.uint())) cursor.skipBuffer();
			return;
		case 8: {
			const length = cursor.uint();
			budget.addStructs(length);
			for (let i = 0; i < length; i++) skipAny(cursor, budget, 0);
			return;
		}
		case 9: // subdocument: guid and options
			cursor.skipBuffer();
			skipAny(cursor, budget, 0);
			return;
		default:
			throw new Overrun(`unknown content type ${ref}`);
	}
}

/** A struct id: client and clock. */
function skipId(cursor: Cursor): void {
	cursor.uint();
	cursor.uint();
}

function skipStruct(cursor: Cursor, budget: Budget): void {
	const info = cursor.byte();
	const ref = info & BITS5;
	if (ref === 0 || ref === 10) {
		// GC and Skip: a length only
		cursor.uint();
		return;
	}
	if (info & BIT8) skipId(cursor); // origin
	if (info & BIT7) skipId(cursor); // right origin
	if ((info & (BIT7 | BIT8)) === 0) {
		// parent: a root type name or the id of the parent type's item
		if (cursor.uint() === 1) cursor.skipBuffer();
		else skipId(cursor);
		if (info & BIT6) cursor.skipBuffer(); // parent map key
	}
	skipContent(cursor, budget, ref);
}

/**
 * Walk `update` (a v1 Yjs update) without decoding it and count what decoding would allocate.
 * Returns the counts, or why the update was refused: a limit crossed or a malformed encoding. Costs
 * one pass over the bytes and allocates nothing in proportion to them.
 */
export function scanUpdate(update: Uint8Array, limits: UpdateLimits = {}): UpdateScan {
	const budget = new Budget({ ...DEFAULT_UPDATE_LIMITS, ...limits });
	const cursor = new Cursor(update);
	try {
		const clients = cursor.uint();
		for (let i = 0; i < clients; i++) {
			const structs = cursor.uint();
			budget.addStructs(structs + 1);
			cursor.uint(); // client
			cursor.uint(); // clock
			for (let j = 0; j < structs; j++) skipStruct(cursor, budget);
		}
		const deleteClients = cursor.uint();
		for (let i = 0; i < deleteClients; i++) {
			cursor.uint(); // client
			const ranges = cursor.uint();
			budget.addDeleteRanges(ranges + 1);
			for (let j = 0; j < ranges; j++) skipId(cursor); // clock and length, read alike
		}
		return {
			ok: true,
			structs: budget.structs,
			deleteRanges: budget.deleteRanges,
			bytesRead: cursor.pos,
		};
	} catch (cause) {
		if (cause instanceof Overrun)
			return { ok: false, reason: cause.message, overLimit: cause.overLimit };
		throw cause;
	}
}
