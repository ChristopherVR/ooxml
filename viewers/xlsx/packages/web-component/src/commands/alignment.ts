// Home > Alignment: vertical and horizontal alignment, orientation, wrap text, indent and the
// merge variants. Checked state reflects the active cell.
import type { HorizontalAlignment, MergeMode, VerticalAlignment } from '@christophervr/xlsx-core';
import { rangesIntersect } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { style } from './font.js';
import { icon } from './icons.js';
import { UNSET, activeStyle, editing, target } from './util.js';

const vertical = (ctx: EditorContext): VerticalAlignment =>
	activeStyle(ctx)?.alignment?.vertical ?? 'bottom';
const horizontal = (ctx: EditorContext): HorizontalAlignment =>
	activeStyle(ctx)?.alignment?.horizontal ?? 'general';

export const ORIENTATIONS: ReadonlyArray<readonly [number, string]> = [
	[45, 'Angle Counterclockwise'],
	[135, 'Angle Clockwise'],
	[255, 'Vertical Text'],
	[90, 'Rotate Text Up'],
	[180, 'Rotate Text Down'],
];

const mergedAtActive = (ctx: EditorContext): boolean => {
	const t = target(ctx);
	return !!t && t.ws.merges.some((m) => rangesIntersect(m, t.range));
};

function mergeCommand(id: string, label: string, mode: MergeMode): Command {
	return editing({
		id,
		label,
		icon: icon('mergeCenter'),
		lock: 'formatCells',
		...(mode === 'center' ? { checked: mergedAtActive } : {}),
		run: (ctx) => {
			const t = target(ctx);
			if (!t) return;
			if (mode === 'center' && mergedAtActive(ctx)) {
				t.session.batch('Unmerge Cells', () => {
					for (const range of t.ranges) t.session.unmerge(t.sheet, range);
				});
				return;
			}
			const lossy = t.ranges.some((r) => {
				let filled = 0;
				for (let row = r.start.row; row <= r.end.row; row++)
					for (const [col, cell] of t.ws.rows.get(row) ?? [])
						if (col >= r.start.col && col <= r.end.col && cell.value !== null && cell.value !== '')
							filled++;
				return mode === 'across' ? false : filled > 1;
			});
			if (lossy)
				ctx.toast(
					ctx.t('Merging cells only keeps the upper-left value and discards other values.'),
					'warning',
				);
			t.session.batch(label, () => {
				for (const range of t.ranges) t.session.merge(t.sheet, range, mode);
			});
		},
	});
}

export function alignmentCommands(): Command[] {
	const valign = (id: string, label: string, value: VerticalAlignment, glyph: string): Command =>
		editing({
			id,
			label,
			icon: icon(glyph),
			lock: 'formatCells',
			checked: (ctx) => vertical(ctx) === value,
			run: (ctx) => style(ctx, { alignment: { vertical: value === 'bottom' ? UNSET : value } }),
		});
	const halign = (id: string, label: string, value: HorizontalAlignment, glyph: string): Command =>
		editing({
			id,
			label,
			icon: icon(glyph),
			lock: 'formatCells',
			checked: (ctx) => horizontal(ctx) === value,
			run: (ctx) =>
				style(ctx, { alignment: { horizontal: horizontal(ctx) === value ? UNSET : value } }),
		});
	return [
		valign('home.align-top', 'Top Align', 'top', 'alignTop'),
		valign('home.align-middle', 'Middle Align', 'center', 'alignMiddle'),
		valign('home.align-bottom', 'Bottom Align', 'bottom', 'alignBottom'),
		halign('home.align-left', 'Align Left', 'left', 'alignLeft'),
		halign('home.align-center', 'Center', 'center', 'alignCenter'),
		halign('home.align-right', 'Align Right', 'right', 'alignRight'),
		editing({
			id: 'home.orientation',
			label: 'Orientation',
			icon: icon('orientation'),
			lock: 'formatCells',
			checked: (ctx) => (activeStyle(ctx)?.alignment?.textRotation ?? 0) !== 0,
			run: (ctx, arg) => {
				const current = activeStyle(ctx)?.alignment?.textRotation ?? 0;
				const next = typeof arg === 'number' ? arg : 45;
				style(ctx, { alignment: { textRotation: current === next ? UNSET : next } });
			},
		}),
		editing({
			id: 'home.alignment-settings',
			label: 'Format Cell Alignment',
			icon: icon('formatCells'),
			lock: 'formatCells',
			run: (ctx) => void ctx.dialogs.open('format-cells', { tab: 'alignment' }),
		}),
		editing({
			id: 'home.wrap-text',
			label: 'Wrap Text',
			icon: icon('wrapText'),
			lock: 'formatCells',
			checked: (ctx) => !!activeStyle(ctx)?.alignment?.wrapText,
			run: (ctx) =>
				style(ctx, {
					alignment: { wrapText: activeStyle(ctx)?.alignment?.wrapText ? UNSET : true },
				}),
		}),
		...([-1, 1] as const).map((delta) =>
			editing({
				id: delta < 0 ? 'home.indent-decrease' : 'home.indent-increase',
				label: delta < 0 ? 'Decrease Indent' : 'Increase Indent',
				icon: icon(delta < 0 ? 'indentDecrease' : 'indentIncrease'),
				lock: 'formatCells',
				run: (ctx) => {
					const current = activeStyle(ctx)?.alignment?.indent ?? 0;
					const next = Math.max(0, Math.min(250, current + delta));
					const h = horizontal(ctx);
					const keep = h === 'right' || h === 'distributed' || h === 'left';
					style(ctx, {
						alignment: {
							indent: next || UNSET,
							horizontal: next ? (keep ? h : 'left') : h === 'general' ? UNSET : h,
						},
					});
				},
			}),
		),
		mergeCommand('home.merge-center', 'Merge & Center', 'center'),
		mergeCommand('home.merge-across', 'Merge Across', 'across'),
		mergeCommand('home.merge-cells', 'Merge Cells', 'merge'),
		editing({
			id: 'home.unmerge',
			label: 'Unmerge Cells',
			icon: icon('mergeCenter'),
			lock: 'formatCells',
			enabled: mergedAtActive,
			run: (ctx) => {
				const t = target(ctx);
				if (!t) return;
				t.session.batch('Unmerge Cells', () => {
					for (const range of t.ranges) t.session.unmerge(t.sheet, range);
				});
			},
		}),
	];
}
