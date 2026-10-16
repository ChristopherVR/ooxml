import type { XmlElement } from '../../xml/index';
import { getCell } from '../cells';
import type {
	PivotAggregate,
	PivotCaptionOverrides,
	PivotCaptions,
	PivotTable,
	Worksheet,
} from '../model';
import { att, decodeEscapes } from '../read/xml-util';
import { pivotGrid } from './aggregate';
import { ENGLISH_PIVOT_CAPTIONS, pivotCaptions, trimmedCaptions } from './captions';

const ATTRIBUTES = {
	rowHeaderCaption: 'rowLabels',
	colHeaderCaption: 'columnLabels',
	grandTotalCaption: 'grandTotal',
	dataCaption: 'values',
} as const;

/** The captions the definition itself names: the header, grand total and Values captions. */
export function readPivotCaptionAttributes(root: XmlElement): PivotCaptionOverrides {
	const out: PivotCaptionOverrides = {};
	for (const [attribute, key] of Object.entries(ATTRIBUTES)) {
		const text = att(root, attribute);
		if (text !== undefined)
			out[key as (typeof ATTRIBUTES)[keyof typeof ATTRIBUTES]] = decodeEscapes(text);
	}
	return out;
}

/** The default value field names the file's data fields follow (`Summe von {0}`). */
export function aggregateCaptionsOf(pivot: PivotTable): PivotCaptionOverrides['aggregates'] {
	const out: Partial<Record<PivotAggregate, string>> = {};
	for (const measure of [
		{ field: pivot.valueField, aggregate: pivot.aggregate, caption: pivot.valueCaption },
		...(pivot.additionalValues ?? []),
	]) {
		const at = measure.caption.indexOf(measure.field);
		if (at < 0 || measure.caption === measure.field) continue;
		const template = `${measure.caption.slice(0, at)}{0}${measure.caption.slice(at + measure.field.length)}`;
		if (template !== ENGLISH_PIVOT_CAPTIONS.aggregates[measure.aggregate])
			out[measure.aggregate] = template;
	}
	return Object.keys(out).length ? out : undefined;
}

const MARK = '\u0001';
const SENTINELS: PivotCaptions = {
	rowLabels: `${MARK}R`,
	columnLabels: `${MARK}C`,
	grandTotal: `${MARK}G`,
	values: `${MARK}V`,
	all: `${MARK}A`,
	blank: `${MARK}B`,
	subtotal: `{0}${MARK}S`,
	measureTotal: `${MARK}M{0}`,
	functions: Object.fromEntries(
		Object.keys(ENGLISH_PIVOT_CAPTIONS.functions).map((fn) => [fn, `${MARK}F${fn}`]),
	) as Record<PivotAggregate, string>,
	aggregates: ENGLISH_PIVOT_CAPTIONS.aggregates,
};

/** A template for `cached` where `arg` stands for `{0}`, preferring the start or the end. */
function templateOf(cached: string, arg: string): string | undefined {
	if (!arg || !cached.includes(arg)) return undefined;
	if (cached.startsWith(arg)) return `{0}${cached.slice(arg.length)}`;
	if (cached.endsWith(arg)) return `${cached.slice(0, -arg.length)}{0}`;
	const at = cached.indexOf(arg);
	return `${cached.slice(0, at)}{0}${cached.slice(at + arg.length)}`;
}

/**
 * What the file's cached cells call the captions its attributes do not name. The grid is built
 * with a marker in place of each caption; the cell the marker lands on holds the file's word.
 * A word the cells disagree on is left English, so the layout check refuses the report rather
 * than rewriting cells with a guess.
 */
export function inferPivotCaptions(
	pivot: PivotTable,
	sheet: Worksheet,
	named: PivotCaptionOverrides,
): PivotCaptionOverrides | undefined {
	const known = pivotCaptions({ captions: named });
	const probe: PivotCaptions = { ...SENTINELS };
	for (const key of Object.keys(named) as (keyof PivotCaptionOverrides)[])
		if (typeof named[key] === 'string')
			(probe as unknown as Record<string, string>)[key] = known[key as 'values'];
	const grid = pivotGrid({ ...pivot, captions: probe as never });
	const found = new Map<string, string | null>();
	const set = (key: string, value: string | undefined) => {
		if (value === undefined) return;
		const before = found.get(key);
		found.set(key, before === undefined || before === value ? value : null);
	};
	grid.values.forEach((line, r) =>
		line.forEach((expected, c) => {
			if (typeof expected !== 'string' || !expected.includes(MARK)) return;
			const cached = getCell(
				sheet,
				pivot.location.start.row + r,
				pivot.location.start.col + c,
			)?.value;
			if (typeof cached !== 'string') return;
			const kind = /\u0001([RCGVASMFB])(\w*)/.exec(expected);
			if (!kind) return;
			const at = kind.index;
			switch (kind[1]) {
				case 'R':
					return set('rowLabels', cached);
				case 'C':
					return set('columnLabels', cached);
				case 'G':
					return set('grandTotal', cached);
				case 'V':
					return set('values', cached);
				case 'A':
					return set('all', cached);
				case 'B':
					return set('blank', cached);
				case 'S':
					return set('subtotal', templateOf(cached, expected.slice(0, at)));
				case 'M':
					return set('measureTotal', templateOf(cached, expected.slice(at + 2)));
				case 'F': {
					// `East <function>`, or `East <function> of Amount` with several measures.
					const item = expected.slice(0, at - 1);
					const of = / of (.*)$/.exec(expected.slice(at));
					const word = cached.slice(item.length + 1, of ? cached.length - of[0].length : undefined);
					if (cached.startsWith(`${item} `) && (!of || cached.endsWith(of[0])))
						set(`function:${kind[2]}`, word);
				}
			}
		}),
	);
	const out: PivotCaptions = pivotCaptions({ captions: named });
	const functions = { ...out.functions };
	for (const [key, value] of found) {
		if (value === null) continue;
		if (key.startsWith('function:')) functions[key.slice(9) as PivotAggregate] = value;
		else (out as unknown as Record<string, string>)[key] = value;
	}
	out.functions = functions;
	return trimmedCaptions(out);
}
