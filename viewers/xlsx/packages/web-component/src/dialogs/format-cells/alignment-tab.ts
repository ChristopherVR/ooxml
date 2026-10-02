// Format Cells > Alignment: horizontal and vertical alignment, indent, orientation, wrap, shrink,
// merge and text direction. Only the fields the user changed are applied.
import type { Alignment, HorizontalAlignment, VerticalAlignment } from '@christophervr/xlsx-core';
import { rangesIntersect } from '@christophervr/xlsx-core';
import { UNSET } from '../../commands/util.js';
import { checkbox, el, field, fieldset, numberInput, row, select } from '../fields.js';
import type { FormatTab, TabInit } from './types.js';

const HORIZONTAL: ReadonlyArray<readonly [HorizontalAlignment, string]> = [
	['general', 'General'],
	['left', 'Left (Indent)'],
	['center', 'Center'],
	['right', 'Right (Indent)'],
	['fill', 'Fill'],
	['justify', 'Justify'],
	['centerContinuous', 'Center Across Selection'],
	['distributed', 'Distributed (Indent)'],
];

const VERTICAL: ReadonlyArray<readonly [VerticalAlignment, string]> = [
	['top', 'Top'],
	['center', 'Center'],
	['bottom', 'Bottom'],
	['justify', 'Justify'],
	['distributed', 'Distributed'],
];

const DIRECTIONS: ReadonlyArray<readonly [string, string]> = [
	['0', 'Context'],
	['1', 'Left-to-Right'],
	['2', 'Right-to-Left'],
];

/** Excel's textRotation (0-90 up, 91-180 down) as dialog degrees (-90 to 90). */
export const rotationToDegrees = (r: number): number =>
	r > 90 && r <= 180 ? 90 - r : r === 255 ? 0 : r;
export const degreesToRotation = (d: number): number => (d < 0 ? 90 - d : d);

export function alignmentTab(init: TabInit): FormatTab {
	const { ctx } = init;
	const a: Alignment = init.style.alignment ?? {};
	const changed = new Set<keyof Alignment | 'merge'>();
	const mark = (key: keyof Alignment | 'merge') => () => changed.add(key);
	const horizontal = select(ctx, HORIZONTAL, a.horizontal ?? 'general');
	const vertical = select(ctx, VERTICAL, a.vertical ?? 'bottom');
	const indent = numberInput(ctx, a.indent ?? 0, 0, 250);
	const degrees = numberInput(ctx, rotationToDegrees(a.textRotation ?? 0), -90, 90);
	const vertText = checkbox(ctx, 'Vertical text', a.textRotation === 255);
	const wrap = checkbox(ctx, 'Wrap text', !!a.wrapText);
	const shrink = checkbox(ctx, 'Shrink to fit', !!a.shrinkToFit);
	const t = init.target;
	const merged = !!t && t.ws.merges.some((m) => t.ranges.some((r) => rangesIntersect(m, r)));
	const merge = checkbox(ctx, 'Merge cells', merged);
	const direction = select(ctx, DIRECTIONS, String(a.readingOrder ?? 0));
	horizontal.addEventListener('change', mark('horizontal'));
	vertical.addEventListener('change', mark('vertical'));
	indent.addEventListener('input', mark('indent'));
	degrees.addEventListener('input', () => {
		changed.add('textRotation');
		vertText.input.checked = false;
	});
	vertText.input.addEventListener('change', mark('textRotation'));
	wrap.input.addEventListener('change', mark('wrapText'));
	shrink.input.addEventListener('change', mark('shrinkToFit'));
	merge.input.addEventListener('change', mark('merge'));
	direction.addEventListener('change', mark('readingOrder'));
	if (!t) merge.input.disabled = true;
	const panel = el(ctx, 'div');
	panel.append(
		row(
			ctx,
			fieldset(
				ctx,
				'Text alignment',
				field(ctx, 'Horizontal:', horizontal),
				field(ctx, 'Vertical:', vertical),
				field(ctx, 'Indent:', indent),
			),
			fieldset(ctx, 'Orientation', field(ctx, 'Degrees', degrees), vertText.wrapper),
		),
		fieldset(ctx, 'Text control', wrap.wrapper, shrink.wrapper, merge.wrapper),
		fieldset(ctx, 'Right-to-left', field(ctx, 'Text direction:', direction)),
	);
	const patch = (): Partial<Alignment> => {
		const out: Partial<Alignment> = {};
		if (changed.has('horizontal'))
			out.horizontal =
				horizontal.value === 'general' ? UNSET : (horizontal.value as HorizontalAlignment);
		if (changed.has('vertical'))
			out.vertical = vertical.value === 'bottom' ? UNSET : (vertical.value as VerticalAlignment);
		if (changed.has('indent')) {
			const n = Math.max(0, Math.min(250, Math.round(Number(indent.value) || 0)));
			out.indent = n || UNSET;
		}
		if (changed.has('textRotation')) {
			const d = Math.max(-90, Math.min(90, Math.round(Number(degrees.value) || 0)));
			out.textRotation = vertText.input.checked ? 255 : degreesToRotation(d) || UNSET;
		}
		if (changed.has('wrapText')) out.wrapText = wrap.input.checked || UNSET;
		if (changed.has('shrinkToFit')) out.shrinkToFit = shrink.input.checked || UNSET;
		if (changed.has('readingOrder')) out.readingOrder = Number(direction.value) || UNSET;
		return out;
	};
	return {
		id: 'alignment',
		label: 'Alignment',
		panel,
		dirty: () => changed.size > 0,
		patch: () => {
			const alignment = patch();
			return Object.keys(alignment).length ? { alignment } : {};
		},
		apply: (target) => {
			if (!changed.has('merge') || merge.input.checked === merged) return;
			for (const range of target.ranges)
				if (merge.input.checked) target.session.merge(target.sheet, range, 'merge');
				else target.session.unmerge(target.sheet, range);
		},
		focus: () => horizontal.focus(),
	};
}
