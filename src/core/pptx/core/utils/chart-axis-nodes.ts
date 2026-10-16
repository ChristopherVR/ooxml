/**
 * Adds and removes whole axis elements of a chart plot area so they match the
 * model's axes: an axis the model gained (Plot Series on Secondary Axis
 * creates a secondary value axis and its hidden category axis) gets a new
 * `c:valAx` / `c:catAx` / `c:dateAx`, and an axis the model dropped that no
 * chart group references any more is removed.
 *
 * New axes are minimal CT_ValAx / CT_CatAx sequences; the per-axis field
 * writers that run afterwards fill in scaling, number format and the rest.
 *
 * @module utils/chart-axis-nodes
 */
import type { PptxChartAxisFormatting, XmlObject } from '../types';

type GetLocalName = (key: string) => string;

const AXIS_TAGS = ['valAx', 'catAx', 'dateAx', 'serAx'];

function asList(value: unknown): XmlObject[] {
	if (value === undefined) return [];
	return (Array.isArray(value) ? value : [value]).filter(
		(item): item is XmlObject => Boolean(item) && typeof item === 'object',
	);
}

function axisIdOf(node: XmlObject, getLocalName: GetLocalName): string | undefined {
	const key = Object.keys(node).find((k) => getLocalName(k) === 'axId');
	const value = key ? (node[key] as XmlObject | undefined)?.['@_val'] : undefined;
	return value === undefined ? undefined : String(value);
}

/** A minimal axis element for `axis` (CT_ValAx, CT_CatAx or CT_DateAx order). */
export function buildChartAxisNode(axis: PptxChartAxisFormatting): XmlObject {
	const node: XmlObject = {
		'c:axId': { '@_val': String(axis.axisId) },
		'c:scaling': { 'c:orientation': { '@_val': axis.orientation ?? 'minMax' } },
		'c:delete': { '@_val': axis.deleted ? '1' : '0' },
		'c:axPos': { '@_val': axis.axPos ?? (axis.axisType === 'valAx' ? 'r' : 'b') },
	};
	if (axis.axisType === 'valAx' || axis.axisType === 'dateAx') {
		node['c:numFmt'] = {
			'@_formatCode': axis.numFmt?.formatCode ?? 'General',
			'@_sourceLinked': axis.numFmt?.sourceLinked === false ? '0' : '1',
		};
	}
	node['c:majorTickMark'] = { '@_val': axis.majorTickMark ?? 'out' };
	node['c:minorTickMark'] = { '@_val': axis.minorTickMark ?? 'none' };
	node['c:tickLblPos'] = { '@_val': axis.tickLblPos ?? 'nextTo' };
	node['c:crossAx'] = { '@_val': String(axis.crossAxisId ?? 0) };
	node['c:crosses'] = {
		'@_val': typeof axis.crosses === 'string' ? axis.crosses : 'autoZero',
	};
	if (axis.axisType === 'valAx') {
		node['c:crossBetween'] = { '@_val': 'between' };
	} else if (axis.axisType === 'catAx') {
		node['c:auto'] = { '@_val': '1' };
		node['c:lblAlgn'] = { '@_val': 'ctr' };
		node['c:lblOffset'] = { '@_val': '100' };
		node['c:noMultiLvlLbl'] = { '@_val': '0' };
	} else if (axis.axisType === 'dateAx') {
		node['c:auto'] = { '@_val': '1' };
		node['c:lblOffset'] = { '@_val': '100' };
	}
	return node;
}

/**
 * Make the plot area's axis elements match `axes`: append an element for
 * every model axis a chart group references but the plot area lacks (after
 * the last axis, before `c:dTable`,
 * `c:spPr` and `c:extLst`), and drop axis elements that are neither in the
 * model nor referenced by any chart group's `c:axId`. Mutates `plotArea`.
 */
export function syncChartAxisNodes(
	plotArea: XmlObject,
	axes: readonly PptxChartAxisFormatting[] | undefined,
	getLocalName: GetLocalName,
): void {
	if (!axes || axes.length === 0) return;
	const entries = Object.entries(plotArea);
	const present = new Set<string>();
	for (const [key, value] of entries) {
		if (!AXIS_TAGS.includes(getLocalName(key))) continue;
		for (const node of asList(value)) {
			const id = axisIdOf(node, getLocalName);
			if (id) present.add(id);
		}
	}
	const referenced = new Set<string>();
	for (const [key, value] of entries) {
		if (!getLocalName(key).endsWith('Chart')) continue;
		for (const container of asList(value)) {
			const axIdKey = Object.keys(container).find((k) => getLocalName(k) === 'axId');
			for (const ref of asList(axIdKey ? container[axIdKey] : undefined)) {
				referenced.add(String(ref['@_val']));
			}
		}
	}
	const modelIds = new Set(axes.map((axis) => String(axis.axisId)));
	const missing = axes.filter(
		(axis) =>
			axis.axisId !== undefined &&
			!present.has(String(axis.axisId)) &&
			referenced.has(String(axis.axisId)) &&
			(axis.axisType === 'valAx' || axis.axisType === 'catAx' || axis.axisType === 'dateAx'),
	);
	let changed = missing.length > 0;
	// Remove orphans: axes the model dropped that nothing points at.
	const kept: Array<[string, unknown]> = [];
	for (const [key, value] of entries) {
		if (!AXIS_TAGS.includes(getLocalName(key))) {
			kept.push([key, value]);
			continue;
		}
		const nodes = asList(value);
		const survivors = nodes.filter((node) => {
			const id = axisIdOf(node, getLocalName);
			return !id || modelIds.has(id) || referenced.has(id);
		});
		if (survivors.length !== nodes.length) changed = true;
		if (survivors.length > 0) kept.push([key, survivors.length === 1 ? survivors[0] : survivors]);
	}
	if (!changed) return;
	// Group new axes by tag, appending to an existing same-tag key when present.
	for (const axis of missing) {
		const tag = axis.axisType;
		const node = buildChartAxisNode(axis);
		const existing = kept.find(([key]) => getLocalName(key) === tag);
		if (existing) {
			existing[1] = [...asList(existing[1]), node];
			continue;
		}
		const lastAxis = kept.reduce(
			(at, [key], index) => (AXIS_TAGS.includes(getLocalName(key)) ? index : at),
			-1,
		);
		const firstAfter = kept.findIndex(([key]) =>
			['dTable', 'spPr', 'extLst'].includes(getLocalName(key)),
		);
		const at = lastAxis >= 0 ? lastAxis + 1 : firstAfter >= 0 ? firstAfter : kept.length;
		kept.splice(at, 0, [`c:${tag}`, node]);
	}
	for (const key of Object.keys(plotArea)) delete plotArea[key];
	for (const [key, value] of kept) plotArea[key] = value as XmlObject[string];
}
