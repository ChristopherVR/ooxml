// Reads Excel 2010 sparkline groups from a worksheet's preserved extension list. The XML itself is
// never changed here: the writer keeps saving the preserved extension, so this is a read-only view.
import { NS, children, first, parseXml, type XmlElement } from '../../xml/index';
import type {
	Color,
	SparklineAxisType,
	SparklineEmptyCells,
	SparklineGroup,
	SparklineType,
	Worksheet,
} from '../model';
import { parseRange } from '../address';
import { parseColor } from './style-parts';
import { boolAttr, descendants, numAttr } from './xml-util';

const XM = 'http://schemas.microsoft.com/office/excel/2006/main';

const TYPES: readonly SparklineType[] = ['line', 'column', 'stacked'];
const EMPTY: readonly SparklineEmptyCells[] = ['gap', 'zero', 'span'];
const AXES: readonly SparklineAxisType[] = ['individual', 'group', 'custom'];

const COLORS = [
	'colorSeries',
	'colorNegative',
	'colorAxis',
	'colorMarkers',
	'colorFirst',
	'colorLast',
	'colorHigh',
	'colorLow',
] as const;

function pick<T extends string>(value: string | null, allowed: readonly T[], fallback: T): T {
	return allowed.includes(value as T) ? (value as T) : fallback;
}

const text = (element: XmlElement | undefined): string | undefined => {
	const value = element?.textContent?.trim();
	return value ? value : undefined;
};

/** Reads one `x14:sparklineGroup`. */
export function readSparklineGroup(node: XmlElement): SparklineGroup {
	const attr = (name: string) => (node.hasAttribute(name) ? node.getAttribute(name) : null);
	const group: SparklineGroup = {
		type: pick(attr('type'), TYPES, 'line'),
		lineWeight: numAttr(node, 'lineWeight') ?? 0.75,
		dateAxis: boolAttr(node, 'dateAxis', false),
		markers: boolAttr(node, 'markers', false),
		high: boolAttr(node, 'high', false),
		low: boolAttr(node, 'low', false),
		first: boolAttr(node, 'first', false),
		last: boolAttr(node, 'last', false),
		negative: boolAttr(node, 'negative', false),
		displayEmptyCellsAs: pick(attr('displayEmptyCellsAs'), EMPTY, 'zero'),
		displayXAxis: boolAttr(node, 'displayXAxis', false),
		displayHidden: boolAttr(node, 'displayHidden', false),
		rightToLeft: boolAttr(node, 'rightToLeft', false),
		minAxisType: pick(attr('minAxisType'), AXES, 'individual'),
		maxAxisType: pick(attr('maxAxisType'), AXES, 'individual'),
		sparklines: [],
	};
	const manualMin = numAttr(node, 'manualMin');
	const manualMax = numAttr(node, 'manualMax');
	if (manualMin !== undefined) group.manualMin = manualMin;
	if (manualMax !== undefined) group.manualMax = manualMax;
	for (const name of COLORS) {
		const color: Color | undefined = parseColor(first(node, name, NS.x14));
		if (color) group[name] = color;
	}
	const dates = text(first(node, 'f', XM));
	if (dates) group.dateFormula = dates;
	for (const line of children(first(node, 'sparklines', NS.x14) ?? node, 'sparkline', NS.x14)) {
		const host = parseRange(text(first(line, 'sqref', XM)) ?? '');
		if (!host) continue;
		const formula = text(first(line, 'f', XM));
		group.sparklines.push(formula ? { formula, host } : { host });
	}
	return group;
}

/** Every sparkline group in the given `extLst` fragments, in document order. */
export function readSparklineGroups(extLst: readonly string[]): SparklineGroup[] {
	const groups: SparklineGroup[] = [];
	for (const xml of extLst) {
		if (!xml.includes('sparklineGroup')) continue;
		let root: XmlElement;
		try {
			root = parseXml(xml).documentElement;
		} catch {
			continue;
		}
		for (const node of descendants(root, 'sparklineGroup', NS.x14))
			groups.push(readSparklineGroup(node));
	}
	return groups;
}

const cache = new WeakMap<Worksheet, { source: readonly string[]; groups: SparklineGroup[] }>();

/**
 * The sparkline groups of a sheet, read from its preserved extension list and cached until that
 * list changes. Treat the result as read-only: the preserved XML is what a save writes.
 */
export function sheetSparklineGroups(sheet: Worksheet): readonly SparklineGroup[] {
	const source = sheet.preserved.get('extLst') ?? [];
	const hit = cache.get(sheet);
	if (hit && hit.source.length === source.length && hit.source.every((xml, i) => xml === source[i]))
		return hit.groups;
	const groups = readSparklineGroups(source);
	cache.set(sheet, { source: [...source], groups });
	return groups;
}
