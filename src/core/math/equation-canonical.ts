/**
 * Comparing and normalising Office Math trees: runs coalesced as Word writes them, and a canonical
 * form that ignores what does not change an equation (control-character formatting, default-valued
 * properties, property order, how text is split into runs), so the editors can tell whether linear
 * text rebuilds an equation exactly before they let it be edited that way.
 */
import { mergeSiblings } from './index';
import type { OmmlNode } from './index';
import { stripXmlOrderSuffix } from './index';
import { readRunText } from './index';
import { ommlToXml } from './index';

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

/** Applies `visit` to every element below `node`, children first. */
function mapTree(node: OmmlNode, visit: (tag: string, value: OmmlNode) => OmmlNode): OmmlNode {
	return rebuild(
		orderedEntries(node).map(([tag, value]): Entry => {
			if (!isNode(value) || tag.startsWith('@_')) return [tag, value];
			return [tag, visit(tag, mapTree(value, visit))];
		}),
	);
}

const runProps = (run: OmmlNode) => JSON.stringify([run['m:rPr'] ?? null, run['w:rPr'] ?? null]);

/** Adjacent runs with the same formatting become one run, as Word writes them. */
export function coalesceRuns(node: OmmlNode): OmmlNode {
	return mapTree(node, (_tag, value) => {
		const entries = orderedEntries(value);
		const out: Entry[] = [];
		for (const entry of entries) {
			const last = out[out.length - 1];
			const [tag, item] = entry;
			if (tag === 'm:r' && last?.[0] === 'm:r' && isNode(item) && isNode(last[1])) {
				const previous = last[1];
				if (runProps(previous) === runProps(item)) {
					last[1] = { ...previous, 'm:t': readRunText(previous) + readRunText(item) };
					continue;
				}
			}
			out.push([tag, item]);
		}
		return out.length === entries.length ? value : rebuild(out);
	});
}

/** On/off properties: present means on, so `m:val="1"` says nothing more. */
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
	'm:show',
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
/** Property values that are the schema default where they are written. */
const DEFAULTS: Record<string, Record<string, string>> = {
	'm:naryPr': { 'm:chr': '∫' },
	'm:dPr': { 'm:begChr': '(', 'm:endChr': ')', 'm:sepChr': '|' },
	'm:accPr': { 'm:chr': '̂' },
	'm:fPr': { 'm:type': 'bar' },
};

function canonicalProps(tag: string, value: OmmlNode): OmmlNode | undefined {
	const out: OmmlNode = {};
	for (const [key, item] of orderedEntries(value).sort(([a], [b]) => a.localeCompare(b))) {
		if (key === 'm:ctrlPr' || key === 'w:rPr') continue;
		const node = isNode(item) ? item : {};
		const val = node['@_val'];
		if (val !== undefined && DEFAULTS[tag]?.[key] === val) continue;
		if (FLAGS.has(key)) {
			if (val === '0' || val === 'off' || val === 'false') continue;
			out[key] = {};
		} else out[key] = item;
	}
	return Object.keys(out).length ? out : undefined;
}

/**
 * The canonical form of an `m:oMath` content tree: formatting of control characters and of
 * runs (`w:rPr`) dropped, default properties dropped, runs coalesced.
 */
export function canonicalMath(math: OmmlNode): string {
	const stripped = mapTree(math, (tag, value) => {
		if (tag === 'm:t') return { '#text': readRunText({ 'm:t': value }) };
		const entries = orderedEntries(value).flatMap(([key, item]): Entry[] => {
			if (key === 'm:ctrlPr' || key === 'w:rPr' || key === '@_space') return [];
			if (key.endsWith('Pr') && isNode(item)) {
				const props = canonicalProps(key, item);
				return props ? [[key, props]] : [];
			}
			return [[key, item]];
		});
		return rebuild(entries);
	});
	return ommlToXml({ 'm:oMath': coalesceRuns(stripped) });
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

/** `math` with `format` as every run's `w:rPr`. */
export function withRunFormat(math: OmmlNode, format: OmmlNode | undefined): OmmlNode {
	if (!format) return math;
	return mapTree(math, (tag, value) => {
		if (tag !== 'm:r') return value;
		const entries = orderedEntries(value).filter(([key]) => key !== 'w:rPr');
		const at = entries.findIndex(([key]) => key !== 'm:rPr');
		entries.splice(at < 0 ? entries.length : at, 0, ['w:rPr', format]);
		return rebuild(entries);
	});
}
