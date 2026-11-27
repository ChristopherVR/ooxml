/**
 * Comparing and normalising Office Math trees: runs coalesced as Word writes them, and a canonical
 * form that ignores what does not change an equation (default-valued properties, property order,
 * how plain text is split into runs, and the math font Word gives every run and control
 * character), so the editors can tell whether linear text rebuilds an equation exactly before they
 * let it be edited that way. Anything else in the tree takes part in the comparison.
 */
import { mergeSiblings, readRunText, stripXmlOrderSuffix, ommlToXml } from './index';
import type { OmmlNode } from './index';

type Entry = [string, unknown];

/** The children of `node` in document order, with plain tag names. */
export function orderedEntries(node: OmmlNode): Entry[] {
	const out: Entry[] = [];
	for (const key of Object.keys(node)) {
		const value = node[key];
		const tag = key.startsWith('@_') || key === '#text' ? key : stripXmlOrderSuffix(key);
		for (const item of Array.isArray(value) ? value : [value]) out.push([tag, item]);
	}
	return out;
}

const rebuild = (entries: Entry[]): OmmlNode =>
	mergeSiblings(entries.map(([tag, value]) => ({ [tag]: value }) as OmmlNode));

const isNode = (value: unknown): value is OmmlNode =>
	Boolean(value) && typeof value === 'object' && !Array.isArray(value);

/** Applies `visit` to every element of `node` and below, children first, `node` itself last. */
function mapTree(
	node: OmmlNode,
	visit: (tag: string, value: OmmlNode) => OmmlNode,
	tag = '',
): OmmlNode {
	const mapped = rebuild(
		orderedEntries(node).map(([child, value]): Entry => {
			if (!isNode(value) || child.startsWith('@_')) return [child, value];
			return [child, mapTree(value, visit, child)];
		}),
	);
	return visit(tag, mapped);
}

/** A run holding only text and run properties, which can join its neighbour. */
function plainRun(run: unknown): run is OmmlNode {
	return (
		isNode(run) &&
		Object.keys(run).every((key) => ['m:rPr', 'w:rPr', 'm:t'].includes(stripXmlOrderSuffix(key)))
	);
}

const runProps = (run: OmmlNode) => JSON.stringify([run['m:rPr'] ?? null, run['w:rPr'] ?? null]);

function mergeRuns(value: OmmlNode): OmmlNode {
	const entries = orderedEntries(value);
	const out: Entry[] = [];
	for (const [tag, item] of entries) {
		const last = out[out.length - 1];
		if (tag === 'm:r' && last?.[0] === 'm:r' && plainRun(item) && plainRun(last[1])) {
			if (runProps(last[1]) === runProps(item)) {
				last[1] = { ...last[1], 'm:t': readRunText(last[1]) + readRunText(item) };
				continue;
			}
		}
		out.push([tag, item]);
	}
	return out.length === entries.length ? value : rebuild(out);
}

/**
 * Adjacent plain runs with the same formatting become one run, as Word writes them, at every
 * level including the top. Runs holding anything else (a symbol, a break) are left alone.
 */
export function coalesceRuns(node: OmmlNode): OmmlNode {
	return mapTree(node, (_tag, value) => mergeRuns(value));
}

/** On/off properties that are off when absent: `m:val="1"` says nothing more than presence. */
const FLAGS = new Set([
	'm:nor',
	'm:lit',
	'm:degHide',
	'm:subHide',
	'm:supHide',
	'm:grow',
	'm:plcHide',
	'm:noBreak',
	'm:opEmu',
	'm:diff',
	'm:aln',
	'm:zeroWid',
	'm:zeroAsc',
	'm:zeroDesc',
	'm:transp',
	'm:maxDist',
	'm:objDist',
	'm:alnScr',
	'm:hideTop',
	'm:hideBot',
	'm:hideLeft',
	'm:hideRight',
	'm:strikeH',
	'm:strikeV',
	'm:strikeBLTR',
	'm:strikeTLBR',
]);
/** On/off properties that are on when absent (a phantom is shown unless `m:show` is off). */
const FLAGS_ON = new Set(['m:show']);
const OFF = new Set(['0', 'off', 'false']);
/** Property values that are the schema default where they are written. */
const DEFAULTS: Record<string, Record<string, string>> = {
	'm:naryPr': { 'm:chr': '∫' },
	'm:dPr': { 'm:begChr': '(', 'm:endChr': ')', 'm:sepChr': '|' },
	'm:accPr': { 'm:chr': '̂' },
	'm:fPr': { 'm:type': 'bar' },
};

/** A matrix column spec of one group of centred columns: what a matrix shows without one. */
function centredColumns(mcs: unknown): boolean {
	const groups = isNode(mcs) ? orderedEntries(mcs) : [];
	if (groups.length !== 1 || groups[0]![0] !== 'm:mc' || !isNode(groups[0]![1])) return false;
	const props = groups[0]![1]['m:mcPr'];
	const jc = isNode(props) && isNode(props['m:mcJc']) ? props['m:mcJc']['@_val'] : undefined;
	return jc === undefined || jc === 'center';
}

function canonicalProps(tag: string, value: OmmlNode): OmmlNode | undefined {
	const out: OmmlNode = {};
	for (const [key, item] of orderedEntries(value).sort(([a], [b]) => a.localeCompare(b))) {
		if (key === 'm:ctrlPr') continue;
		if (tag === 'm:mPr' && key === 'm:mcs' && centredColumns(item)) continue;
		const node = isNode(item) ? item : {};
		const val = node['@_val'];
		if (val !== undefined && DEFAULTS[tag]?.[key] === val) continue;
		const off = typeof val === 'string' && OFF.has(val);
		if (FLAGS.has(key)) {
			if (!off) out[key] = {};
		} else if (FLAGS_ON.has(key)) {
			if (off) out[key] = { '@_val': '0' };
		} else out[key] = item;
	}
	return Object.keys(out).length ? out : undefined;
}

/**
 * The canonical form of an `m:oMath` content tree: control-character formatting (checked apart,
 * see `controlFormat`) and the shared run format (`sharedRunFormat`) dropped, default properties
 * dropped, plain runs coalesced; everything else is compared as it is.
 */
export function canonicalMath(math: OmmlNode): string {
	const stripped = mapTree(math, (tag, value) => {
		if (tag === 'm:t') return { '#text': readRunText({ 'm:t': value }) };
		const entries = orderedEntries(value).flatMap(([key, item]): Entry[] => {
			if (key === 'm:ctrlPr' || (tag === 'm:r' && key === 'w:rPr') || key === '@_space') return [];
			if (key.endsWith('Pr') && key.startsWith('m:') && isNode(item)) {
				const props = canonicalProps(key, item);
				return props ? [[key, props]] : [];
			}
			return [[key, item]];
		});
		return mergeRuns(rebuild(entries));
	});
	return ommlToXml({ 'm:oMath': stripped });
}

/**
 * The `w:rPr` every run of `math` shares (Word gives each run its math font), `undefined` when no
 * run has one, or `'mixed'` when runs are formatted differently.
 */
export function sharedRunFormat(math: OmmlNode): OmmlNode | 'mixed' | undefined {
	const seen = new Set<string>();
	let format: OmmlNode | undefined;
	let runs = 0;
	mapTree(math, (tag, value) => {
		if (tag === 'm:r') {
			runs++;
			const props = value['w:rPr'];
			if (isNode(props)) format = props;
			seen.add(JSON.stringify(props ?? null));
		}
		return value;
	});
	if (seen.size > 1) return 'mixed';
	return runs && format ? format : undefined;
}

/** Formatting a control character may carry and still be rewritten: the math font and italic. */
const CONTROL_KEYS = new Set(['w:rFonts', 'w:i', 'w:iCs']);

/**
 * The formatting of the equation's control characters (`m:ctrlPr`), when every one of them holds
 * only the math font (and italic) and they agree on the font; `'other'` when any holds more (a
 * colour, a tracked change), so the equation is not rewritten.
 */
export function controlFormat(math: OmmlNode): OmmlNode | 'other' | undefined {
	let format: OmmlNode | undefined;
	let fonts: string | undefined;
	let other = false;
	mapTree(math, (tag, value) => {
		if (tag !== 'm:ctrlPr') return value;
		const entries = orderedEntries(value);
		const props = entries.length === 1 && entries[0]![0] === 'w:rPr' ? entries[0]![1] : undefined;
		if (entries.length && !isNode(props)) other = true;
		if (!isNode(props)) return value;
		const keys = orderedEntries(props).map(([key]) => key);
		if (keys.some((key) => !CONTROL_KEYS.has(key))) other = true;
		// The math font is the Latin font; Word varies the complex-script one between controls.
		const fontsNode = isNode(props['w:rFonts']) ? props['w:rFonts'] : {};
		const font = JSON.stringify([fontsNode['@_ascii'] ?? null, fontsNode['@_hAnsi'] ?? null]);
		if (fonts !== undefined && font !== fonts) other = true;
		fonts = font;
		if (props['w:i'] !== undefined) format = props;
		else format ??= props;
		return value;
	});
	return other ? 'other' : format;
}

/** Whether every element of `node` is Office Math or WordprocessingML (no other namespace). */
export function onlyMathAndWord(node: OmmlNode): boolean {
	return orderedEntries(node).every(
		([key, value]) =>
			key.startsWith('@_') ||
			key === '#text' ||
			(/^[mw]:/.test(key) && (!isNode(value) || onlyMathAndWord(value))),
	);
}

/** Elements whose properties Word gives a control-character format (`m:ctrlPr`). */
const STRUCTURES = new Set([
	'm:acc',
	'm:bar',
	'm:borderBox',
	'm:box',
	'm:d',
	'm:eqArr',
	'm:f',
	'm:func',
	'm:groupChr',
	'm:limLow',
	'm:limUpp',
	'm:m',
	'm:nary',
	'm:phant',
	'm:rad',
	'm:sPre',
	'm:sSub',
	'm:sSubSup',
	'm:sSup',
]);

/** `math` with `format` as every run's `w:rPr` and `control` in every structure's properties. */
export function withRunFormat(
	math: OmmlNode,
	format: OmmlNode | undefined,
	control?: OmmlNode,
): OmmlNode {
	if (!format && !control) return math;
	return mapTree(math, (tag, value) => {
		if (tag === 'm:r' && format) {
			const entries = orderedEntries(value).filter(([key]) => key !== 'w:rPr');
			const at = entries.findIndex(([key]) => key !== 'm:rPr');
			entries.splice(at < 0 ? entries.length : at, 0, ['w:rPr', format]);
			return rebuild(entries);
		}
		if (control && STRUCTURES.has(tag)) {
			const pr = `${tag}Pr`;
			const props = isNode(value[pr]) ? value[pr] : {};
			return { ...value, [pr]: { ...props, 'm:ctrlPr': { 'w:rPr': control } } } as OmmlNode;
		}
		return value;
	});
}
