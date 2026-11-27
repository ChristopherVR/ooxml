// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// LISTNUM (ECMA-376 Part 1, 17.16.5.31): a number counted with a list, inside a paragraph. The
// rules follow Word's output (`__fixtures__/computed-fields/list-number-*.docx`):
// - In a numbered paragraph, a LISTNUM without a name counts with the paragraph's own list, at
//   its level, and moves the list on (the next item's number is one higher). Elsewhere it counts
//   with the last list a numbered paragraph or a named LISTNUM used, else NumberDefault.
// - NumberDefault, OutlineDefault and LegalDefault are one set of counters drawn three ways.
// - Without `\l`, the first LISTNUM of a paragraph takes the level of the last LISTNUM that led
//   a paragraph in the same list (one with only spaces and LISTNUMs before it), else level 1; a
//   later one is one level deeper than the one before when the paragraph starts with a LISTNUM,
//   else at the same level. `\l` is clamped to 1-9; `\s` sets the count.
import { OUTLINE_NUMBERING } from './outline-numbering';
import { fieldResultSpans, type FieldResultSpan } from './field-result-spans';
import { fieldName } from './field-runs';
import { parseFieldInstruction } from './field-format';
import type { Paragraph } from './model';
import { NumberingCounters, substituteLvlText } from './numbering-counters';
import type { NumberingCatalog, NumberingLevelDefinition } from './numbering-model';
import { displayListLabel, formatListNumber } from './numbering-format';
import { resolveNumberingLevel } from './numbering-parse';

/** The built-in lists a LISTNUM can name (any case), drawn from one set of counters. */
const DEFAULT_LISTS = ['NumberDefault', 'OutlineDefault', 'LegalDefault'] as const;
type DefaultList = (typeof DEFAULT_LISTS)[number];

/** The counters of the built-in lists, kept apart from the document's numbering. */
const DEFAULT_ID = '\u0000listnum';

/** Level `n` (0-8) of the built-in lists: prefix, number format and suffix. */
const FORMATS: Record<Exclude<DefaultList, 'LegalDefault'>, readonly [string, string, string][]> = {
	NumberDefault: [
		['', 'decimal', ')'],
		['', 'lowerLetter', ')'],
		['', 'lowerRoman', ')'],
		['(', 'decimal', ')'],
		['(', 'lowerLetter', ')'],
		['(', 'lowerRoman', ')'],
		['', 'decimal', '.'],
		['', 'lowerLetter', '.'],
		['', 'lowerRoman', '.'],
	],
	OutlineDefault: OUTLINE_NUMBERING,
};

const level = (index: number): NumberingLevelDefinition => ({
	level: index,
	start: 1,
	numFmt: 'decimal',
	lvlText: '',
});

/** A numbering catalog holding only the built-in lists' counters, for `NumberingCounters`. */
const DEFAULT_CATALOG: NumberingCatalog = {
	abstractNums: {
		[DEFAULT_ID]: {
			id: DEFAULT_ID,
			levels: Object.fromEntries([...Array(9).keys()].map((i) => [i, level(i)])),
		},
	},
	nums: { [DEFAULT_ID]: { id: DEFAULT_ID, abstractNumId: DEFAULT_ID } },
	warnings: [],
};

/** The list a LISTNUM counts with: a built-in one by its name, or a numbering instance. */
type ListTarget = { kind: 'default'; name: DefaultList } | { kind: 'num'; numId: string };

const keyOf = (target: ListTarget) =>
	target.kind === 'default' ? target.name : `#${target.numId}`;

const isListNum = (field: FieldResultSpan) => fieldName(field.instr) === 'LISTNUM';

/** The LISTNUM results of a document, worked out paragraph by paragraph in document order. */
export class ListNumberFields {
	private readonly results = new Map<string, Map<number, string>>();
	private context: ListTarget = { kind: 'default', name: 'NumberDefault' };
	private readonly leadLevels = new Map<string, number>();

	constructor(
		private readonly counters: NumberingCounters,
		private readonly catalog: NumberingCatalog | undefined,
	) {}

	/** What the LISTNUM starting at run `from` of a paragraph (or with its result there) shows. */
	result(paragraphId: string, from: number): string | undefined {
		return this.results.get(paragraphId)?.get(from);
	}

	/**
	 * Counts a paragraph's LISTNUM fields, after its own list number (`numbering`, its list and
	 * zero-based level, when it has one).
	 */
	paragraph(paragraph: Paragraph, numbering?: { numId: string; level: number }): void {
		if (numbering) this.context = { kind: 'num', numId: numbering.numId };
		if (!paragraph.runs.some((run) => /LISTNUM/i.test(run.fieldCode ?? run.field?.instr ?? '')))
			return;
		const fields = fieldResultSpans(paragraph.runs, { resultless: true })
			.filter(isListNum)
			.sort((a, b) => a.from - b.from);
		const blank = (from: number, to: number) =>
			paragraph.runs
				.slice(from, to)
				.every((run) => run.fieldCode !== undefined || !/\S/.test(run.text));
		const leads = fields.length > 0 && blank(0, fields[0]!.from);
		let leading = leads;
		let previous: number | undefined;
		let end = -1;
		const shown = new Map<number, string>();
		for (const field of fields) {
			leading &&= blank(end + 1, field.from);
			end = field.to;
			const parsed = parseFieldInstruction(field.instr, { l: 'arg', s: 'arg' });
			if (!parsed) continue;
			const name = parsed.args[0];
			const named = name && DEFAULT_LISTS.find((item) => item.toLowerCase() === name.toLowerCase());
			if (name && !named) continue;
			const target: ListTarget = named
				? { kind: 'default', name: named }
				: numbering
					? { kind: 'num', numId: numbering.numId }
					: this.context;
			if (named && !numbering) this.context = target;
			const given = parsed.switches.get('l');
			const explicit =
				typeof given === 'string' && /^-?\d+$/.test(given)
					? Math.min(9, Math.max(1, Number(given)))
					: undefined;
			const own =
				explicit ??
				(previous === undefined
					? numbering && !named
						? numbering.level + 1
						: (this.leadLevels.get(keyOf(target)) ?? 1)
					: leads
						? Math.min(9, previous + 1)
						: previous);
			const start = parsed.switches.get('s');
			const text = this.count(
				target,
				own - 1,
				typeof start === 'string' && /^\d+$/.test(start) ? Number(start) : undefined,
			);
			if (text === undefined) continue;
			// Found by the field's first run, or by a run of its saved result.
			for (const at of [field.from, ...field.runs]) shown.set(at, text);
			if (leading) this.leadLevels.set(keyOf(target), own);
			previous = own;
		}
		if (shown.size) this.results.set(paragraph.id, shown);
	}

	/** Advances `target` at zero-based level `wanted` (or sets it to `start`) and draws its number. */
	private count(target: ListTarget, wanted: number, start: number | undefined): string | undefined {
		const numId = target.kind === 'default' ? DEFAULT_ID : target.numId;
		const catalog = target.kind === 'default' ? DEFAULT_CATALOG : this.catalog;
		if (!catalog) return undefined;
		// A list without the level asked for (a single-level list) counts at its deepest one.
		let index = wanted;
		let def = resolveNumberingLevel(catalog, numId, index);
		while (!def && index > 0) def = resolveNumberingLevel(catalog, numId, --index);
		if (!def) return undefined;
		if (start === undefined) this.counters.advance(numId, index, def, catalog);
		else this.counters.set(numId, index, start);
		if (target.kind === 'num')
			return def.numFmt === 'bullet'
				? displayListLabel(def.lvlText, def.markerFormat?.fontFamily)
				: substituteLvlText(def.lvlText, numId, index, catalog, this.counters, Boolean(def.isLgl));
		const value = (at: number) => this.counters.current(DEFAULT_ID, at, 1);
		if (target.name === 'LegalDefault')
			return `${[...Array(index + 1).keys()].map(value).join('.')}.`;
		const [open, numFmt, close] = FORMATS[target.name][index]!;
		return `${open}${formatListNumber(numFmt, value(index))}${close}`;
	}
}
