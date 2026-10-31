// How each sheet part that is a list of objects maps onto keyed entries (see keyed.ts):
//   comments  `A1` -> { author, text }                   one note per cell, keyed by address
//   replies   minted key -> { c: 'A1', author, text, date }  threaded replies, so two peers
//             replying to one thread both keep their reply
//   cf        minted key -> ConditionalFormat (ranges and all its rules)
//   dv        minted key -> DataValidation
//   links     range -> Hyperlink                         one link per range
//   tables    minted key -> Table (columns, style flags, filters), the whole table as one entry
//   charts    minted key -> ChartObject, last writer wins per chart (anchor, name, `chartSpace`)
// Everything the model carries for an item (preserved extension XML inside a rule or chart, the
// chart's neutral `chartSpace`) travels in the entry verbatim; entries are validated on decode.
import { formatAddress, formatRange, parseAddress, parseRange } from '../address';
import type { CellRange } from '../address';
import type {
	ChartObject,
	Comment,
	ConditionalFormat,
	DataValidation,
	Hyperlink,
	Table,
} from '../model';
import type { PartSpec } from './keyed';
import { isRecord } from './schema';

type Reply = NonNullable<Comment['replies']>[number];

const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T;
const isText = (v: unknown, max = 32_767): v is string => typeof v === 'string' && v.length <= max;
const isIndex = (v: unknown, max: number): v is number =>
	typeof v === 'number' && Number.isInteger(v) && v >= 0 && v <= max;

function isRange(raw: unknown): raw is CellRange {
	if (!isRecord(raw) || !isRecord(raw.start) || !isRecord(raw.end)) return false;
	return (
		isIndex(raw.start.row, 1_048_575) &&
		isIndex(raw.start.col, 16_383) &&
		isIndex(raw.end.row, 1_048_575) &&
		isIndex(raw.end.col, 16_383)
	);
}
const rangeList = (raw: unknown): boolean =>
	Array.isArray(raw) && raw.length <= 4096 && raw.every(isRange);

export const commentSpec: PartSpec<Comment> = {
	encode: (c) => ({
		a: formatAddress(c.address),
		author: c.author,
		text: c.text,
		...(c.date !== undefined ? { date: c.date } : {}),
	}),
	decode: (raw) => {
		if (!isRecord(raw) || !isText(raw.a, 12) || !isText(raw.author, 1024) || !isText(raw.text))
			return undefined;
		if (raw.date !== undefined && !isText(raw.date, 64)) return undefined;
		const address = parseAddress(raw.a);
		return address
			? {
					address,
					author: raw.author,
					text: raw.text,
					...(raw.date !== undefined ? { date: raw.date } : {}),
				}
			: undefined;
	},
	fixedKey: (e) => String((e as { a?: unknown }).a),
	adopt: (decoded, local) => {
		if (local?.replies) decoded.replies = local.replies;
	},
};

/** Replies need to know their comment: `owner` maps a reply to its comment's address. */
export function replySpec(
	owner: Map<object, string>,
	decodedOwner: Map<object, string>,
): PartSpec<Reply> {
	return {
		encode: (r) => {
			const c = owner.get(r) ?? decodedOwner.get(r);
			if (c === undefined) return undefined;
			const out: Record<string, unknown> = { c, author: r.author, text: r.text };
			if (r.date !== undefined) out.date = r.date;
			return out;
		},
		decode: (raw) => {
			if (!isRecord(raw) || !isText(raw.c, 12) || !isText(raw.author, 1024) || !isText(raw.text))
				return undefined;
			if (raw.date !== undefined && !isText(raw.date, 64)) return undefined;
			const reply: Reply = { author: raw.author, text: raw.text };
			if (raw.date !== undefined) reply.date = raw.date;
			decodedOwner.set(reply, raw.c);
			return reply;
		},
	};
}

export const conditionalSpec: PartSpec<ConditionalFormat> = {
	encode: (cf) => clone(cf),
	decode: (raw) =>
		isRecord(raw) && rangeList(raw.ranges) && Array.isArray(raw.rules) && raw.rules.length <= 1024
			? (clone(raw) as unknown as ConditionalFormat)
			: undefined,
};

export const validationSpec: PartSpec<DataValidation> = {
	encode: (dv) => clone(dv),
	decode: (raw) =>
		isRecord(raw) && rangeList(raw.ranges) && isText(raw.type, 32)
			? (clone(raw) as unknown as DataValidation)
			: undefined,
};

export const linkSpec: PartSpec<Hyperlink> = {
	encode: (link) => {
		const { range, ...rest } = clone(link);
		return { r: formatRange(range), ...rest };
	},
	decode: (raw) => {
		if (!isRecord(raw) || !isText(raw.r, 24)) return undefined;
		const range = parseRange(raw.r);
		if (!range) return undefined;
		const { r: _r, ...rest } = clone(raw);
		return { ...(rest as object), range } as Hyperlink;
	},
	fixedKey: (e) => String((e as { r?: unknown }).r),
};

/** A copy without the part name, which belongs to the package of the peer that holds it. */
function withoutPart<T extends { partName?: string | undefined }>(item: T): T {
	const out = clone(item);
	delete out.partName;
	return out;
}

export const tableSpec: PartSpec<Table> = {
	encode: (table) => withoutPart(table),
	decode: (raw) =>
		isRecord(raw) &&
		isRange(raw.range) &&
		isText(raw.name, 255) &&
		isText(raw.displayName, 255) &&
		Array.isArray(raw.columns) &&
		raw.columns.length <= 16_384
			? (clone(raw) as unknown as Table)
			: undefined,
	natural: (e) => String((e as { displayName?: unknown }).displayName),
	adopt: (decoded, local) => {
		if (local?.partName !== undefined) decoded.partName = local.partName;
	},
};

export const chartSpec: PartSpec<ChartObject> = {
	encode: (chart) => withoutPart(chart),
	decode: (raw) =>
		isRecord(raw) && raw.kind === 'chart' && isRecord(raw.anchor) && isRecord(raw.anchor.from)
			? (clone(raw) as unknown as ChartObject)
			: undefined,
	natural: (e) => {
		const name = (e as { name?: unknown }).name;
		return typeof name === 'string' ? name : undefined;
	},
	adopt: (decoded, local) => {
		if (local?.partName !== undefined) decoded.partName = local.partName;
	},
};
