import { NS, buildXml, children, first, parseXml } from '../../xml/index';
import { formatRange, parseRange, rangesIntersect, type CellRange } from '../address';
import type { SparklineGroup, Worksheet } from '../model';
import { readSparklineGroup, sheetSparklineGroups } from '../read/sparklines';
import { el, escapeText } from '../write/xml-out';
import { type EditContext, sheetAt } from './context';
import { checkedSparklineRange, newSparklines } from './sparkline-ranges';
import {
	XM,
	appendSparklineXml,
	cleanSparklinePatch,
	editSparklineXml,
	patchSparklineNode,
	type SparklineGroupPatch,
} from './sparkline-xml';
export type { SparklineGroupPatch } from './sparkline-xml';

/** Groups containing a sparkline host in the selection. Settings apply to each whole group. */
export function selectedSparklineGroups(
	sheet: Worksheet,
	ranges: readonly CellRange[],
): readonly SparklineGroup[] {
	return sheetSparklineGroups(sheet).filter((group) =>
		group.sparklines.some((line) => ranges.some((range) => rangesIntersect(range, line.host))),
	);
}

function validateSettings(group: SparklineGroup): void {
	for (const key of [
		'markers',
		'high',
		'low',
		'first',
		'last',
		'negative',
		'displayXAxis',
		'displayHidden',
		'rightToLeft',
	] as const)
		if (typeof group[key] !== 'boolean') throw new Error('The sparkline settings are not valid.');
	if (
		!['line', 'column', 'stacked'].includes(group.type) ||
		!['gap', 'zero', 'span'].includes(group.displayEmptyCellsAs) ||
		!['individual', 'group', 'custom'].includes(group.minAxisType) ||
		!['individual', 'group', 'custom'].includes(group.maxAxisType)
	)
		throw new Error('The sparkline settings are not valid.');
	if (!Number.isFinite(group.lineWeight) || group.lineWeight < 0 || group.lineWeight > 1584)
		throw new Error('The sparkline settings are not valid.');
	if (
		(group.minAxisType === 'custom' && !Number.isFinite(group.manualMin)) ||
		(group.maxAxisType === 'custom' && !Number.isFinite(group.manualMax)) ||
		(group.minAxisType === 'custom' &&
			group.maxAxisType === 'custom' &&
			group.manualMin! > group.manualMax!)
	)
		throw new Error('The sparkline axis limits are not valid.');
	if (
		(group.manualMin !== undefined && group.minAxisType !== 'custom') ||
		(group.manualMax !== undefined && group.maxAxisType !== 'custom')
	)
		throw new Error('The sparkline axis limits are not valid.');
	for (const key of [
		'colorSeries',
		'colorNegative',
		'colorAxis',
		'colorMarkers',
		'colorFirst',
		'colorLast',
		'colorHigh',
		'colorLow',
	] as const) {
		const color = group[key];
		if (
			color !== undefined &&
			(color === null ||
				typeof color !== 'object' ||
				Array.isArray(color) ||
				![color.rgb, color.theme, color.indexed].some((value) => value !== undefined))
		)
			throw new Error('The sparkline settings are not valid.');
		if (
			color &&
			(color.auto ||
				(color.rgb !== undefined && !/^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(color.rgb)) ||
				(color.theme !== undefined &&
					(!Number.isInteger(color.theme) || color.theme < 0 || color.theme > 11)) ||
				(color.indexed !== undefined &&
					(!Number.isInteger(color.indexed) || color.indexed < 0 || color.indexed > 65)) ||
				(color.tint !== undefined &&
					(!Number.isFinite(color.tint) || color.tint < -1 || color.tint > 1)))
		)
			throw new Error('The sparkline settings are not valid.');
	}
}

function validatePatch(group: SparklineGroup, patch: SparklineGroupPatch): void {
	const next = { ...group, ...patch };
	if (patch.minAxisType && patch.minAxisType !== 'custom' && patch.manualMin === undefined)
		delete next.manualMin;
	if (patch.maxAxisType && patch.maxAxisType !== 'custom' && patch.manualMax === undefined)
		delete next.manualMax;
	validateSettings(next);
}

function save(
	ctx: EditContext,
	s: number,
	next: string[],
	label: string,
	ranges: CellRange[],
): void {
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		label,
		'annotations',
		[{ kind: 'parts', sheet: s, parts: ['preserved'] }],
		() => sheet.preserved.set('extLst', next),
		{ sheet: s, ranges },
	);
}

/** Creates one group without changing host cell values or other preserved extensions. */
export function createSparklineGroup(
	ctx: EditContext,
	s: number,
	source: string,
	locations: CellRange,
	patch: SparklineGroupPatch = {},
): void {
	patch = cleanSparklinePatch(patch);
	locations = checkedSparklineRange(locations);
	const sheet = sheetAt(ctx.workbook, s);
	const lines = newSparklines(ctx.workbook, s, source, locations);
	if (selectedSparklineGroups(sheet, [locations]).length)
		throw new Error('The location range already contains sparklines.');
	if (sheet.merges.some((merge) => lines.some((line) => rangesIntersect(merge, line.host))))
		throw new Error('Sparklines cannot be created in merged cells.');
	const body = lines
		.map(
			(line) =>
				`<x14:sparkline>${el('xm:f', {}, escapeText(line.formula!))}${el('xm:sqref', {}, formatRange(line.host))}</x14:sparkline>`,
		)
		.join('');
	const colors =
		'<x14:colorSeries theme="4"/><x14:colorNegative rgb="FFD00000"/><x14:colorAxis rgb="FF000000"/>' +
		['Markers', 'First', 'Last', 'High', 'Low']
			.map((name) => `<x14:color${name} rgb="FFD00000"/>`)
			.join('');
	const doc = parseXml(
		`<x14:sparklineGroup xmlns:x14="${NS.x14}" xmlns:xm="${XM}">${colors}<x14:sparklines>${body}</x14:sparklines></x14:sparklineGroup>`,
	);
	const node = doc.documentElement;
	validatePatch(readSparklineGroup(node), patch);
	patchSparklineNode(node, { displayEmptyCellsAs: 'gap', ...patch });
	// Validate the patched XML through the existing reader before it enters history.
	validateSettings({ ...readSparklineGroup(node), ...patch });
	save(ctx, s, appendSparklineXml(sheet, buildXml(node)), 'Create sparklines', [locations]);
}

/** Changes settings for each entire selected group, retaining unknown imported fields. */
export function setSparklineGroups(
	ctx: EditContext,
	s: number,
	ranges: CellRange[],
	patch: SparklineGroupPatch,
): void {
	patch = cleanSparklinePatch(patch);
	const sheet = sheetAt(ctx.workbook, s);
	ranges = ranges.map(checkedSparklineRange);
	const selected = selectedSparklineGroups(sheet, ranges);
	if (!selected.length) return;
	for (const group of selected) validatePatch(group, patch);
	const next = editSparklineXml(sheet, (node, group) => {
		if (!group.sparklines.some((line) => ranges.some((range) => rangesIntersect(range, line.host))))
			return false;
		patchSparklineNode(node, patch);
		return true;
	});
	save(ctx, s, next, 'Format sparklines', ranges);
}

/** Clears selected hosts, or all sparklines in their groups when wholeGroups is true. */
export function deleteSparklines(
	ctx: EditContext,
	s: number,
	ranges: CellRange[],
	wholeGroups = false,
): void {
	ranges = ranges.map(checkedSparklineRange);
	const sheet = sheetAt(ctx.workbook, s);
	if (!selectedSparklineGroups(sheet, ranges).length) return;
	const next = editSparklineXml(sheet, (node, group) => {
		if (!group.sparklines.some((line) => ranges.some((range) => rangesIntersect(range, line.host))))
			return false;
		const lines = first(node, 'sparklines', NS.x14)!;
		const nodes = children(lines, 'sparkline', NS.x14);
		for (const line of nodes) {
			const host = parseRange(first(line, 'sqref', XM)?.textContent ?? '');
			if (wholeGroups || (host && ranges.some((range) => rangesIntersect(range, host))))
				line.parentNode?.removeChild(line);
		}
		return true;
	});
	save(ctx, s, next, 'Clear sparklines', ranges);
}
