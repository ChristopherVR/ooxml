// Sparklines > Group / Ungroup. A group shares its settings (type, colours, axis), so grouping
// moves the sparklines of the selected groups into the first one (whose settings win, as Excel
// applies the active sparkline's), and ungrouping gives every sparkline of the selected groups a
// group of its own with a copy of those settings.
import { type CellRange, rangesIntersect } from '../address';
import { NS, children, first, type XmlElement } from '../../xml/index';
import { type EditContext, sheetAt } from './context';
import { editSparklineXml } from './sparkline-xml';
import { checkedSparklineRange } from './sparkline-ranges';

const touches = (group: { sparklines: { host: CellRange }[] }, ranges: CellRange[]): boolean =>
	group.sparklines.some((line) => ranges.some((r) => rangesIntersect(r, line.host)));

/**
 * Merges the groups the ranges touch into one; returns how many groups were merged (0 or 1 when
 * the selection is already a single group, in which case nothing changes).
 */
export function groupSparklines(ctx: EditContext, s: number, ranges: CellRange[]): number {
	ranges = ranges.map(checkedSparklineRange);
	const sheet = sheetAt(ctx.workbook, s);
	let target: XmlElement | undefined;
	let merged = 0;
	// First pass counts the touched groups so a single group is left alone.
	let touched = 0;
	editSparklineXml(sheet, (_node, group) => {
		if (touches(group, ranges)) touched++;
		return false;
	});
	if (touched < 2) return touched;
	const next = editSparklineXml(sheet, (node, group) => {
		if (!touches(group, ranges)) return false;
		const lines = first(node, 'sparklines', NS.x14);
		if (!target) {
			target = node;
			merged = 1;
			return true;
		}
		const into = first(target, 'sparklines', NS.x14);
		if (!lines || !into) return false;
		for (const line of children(lines, 'sparkline', NS.x14)) into.appendChild(line);
		merged++;
		return true;
	});
	ctx.run(
		'Group sparklines',
		'annotations',
		[{ kind: 'parts', sheet: s, parts: ['preserved'] }],
		() => sheet.preserved.set('extLst', next),
		{ sheet: s, ranges },
	);
	return merged;
}

/** Splits every group the ranges touch into one group per sparkline; returns the groups made. */
export function ungroupSparklines(ctx: EditContext, s: number, ranges: CellRange[]): number {
	ranges = ranges.map(checkedSparklineRange);
	const sheet = sheetAt(ctx.workbook, s);
	let made = 0;
	const next = editSparklineXml(sheet, (node, group) => {
		if (!touches(group, ranges)) return false;
		const lines = first(node, 'sparklines', NS.x14);
		const all = lines ? children(lines, 'sparkline', NS.x14) : [];
		if (all.length < 2) return false;
		let anchor = node;
		for (const line of all.slice(1)) {
			const copy = node.cloneNode(true) as XmlElement;
			const holder = first(copy, 'sparklines', NS.x14)!;
			for (const old of children(holder, 'sparkline', NS.x14)) holder.removeChild(old);
			lines!.removeChild(line);
			holder.appendChild(line);
			anchor.parentNode!.insertBefore(copy, anchor.nextSibling);
			anchor = copy;
			made++;
		}
		return true;
	});
	if (!made) return 0;
	ctx.run(
		'Ungroup sparklines',
		'annotations',
		[{ kind: 'parts', sheet: s, parts: ['preserved'] }],
		() => sheet.preserved.set('extLst', next),
		{ sheet: s, ranges },
	);
	return made;
}
