// Home > Font: font name and size, grow / shrink, bold, italic, underline, strikethrough, font and
// fill colours and the border presets. Toggle state reads the active cell's resolved format.
import type { BorderEdge, BorderPreset, Color, StylePatch } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { icon } from './icons.js';
import { UNSET, activeStyle, editing, target, withTarget } from './util.js';

export const FONT_NAMES = [
	'Aptos',
	'Arial',
	'Arial Black',
	'Calibri',
	'Calibri Light',
	'Cambria',
	'Candara',
	'Century Gothic',
	'Comic Sans MS',
	'Consolas',
	'Constantia',
	'Corbel',
	'Courier New',
	'Franklin Gothic Medium',
	'Garamond',
	'Georgia',
	'Impact',
	'Lucida Console',
	'Lucida Sans Unicode',
	'Palatino Linotype',
	'Segoe UI',
	'Tahoma',
	'Times New Roman',
	'Trebuchet MS',
	'Verdana',
] as const;

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72] as const;

/** The next size in Excel's list (grow) or the previous one (shrink). */
export function stepFontSize(size: number, grow: boolean): number {
	if (grow) return FONT_SIZES.find((s) => s > size) ?? Math.min(409, size + 10);
	return [...FONT_SIZES].reverse().find((s) => s < size) ?? Math.max(1, size - 1);
}

export const style = (ctx: EditorContext, patch: StylePatch): void => {
	withTarget(ctx, (t) => t.session.applyStyle(t.sheet, t.ranges, patch));
};

/**
 * Colour command args, as the ribbon's colour split sends them: a `Color` applies it; undefined,
 * null or an `auto` colour clears it (Automatic font colour, No Fill).
 */
export function pickColor(arg: unknown): Color | null {
	if (arg && typeof arg === 'object' && !(arg as Color).auto) return arg as Color;
	return null;
}

const toggle = (
	id: string,
	label: string,
	glyph: string,
	shortcut: string,
	read: (ctx: EditorContext) => boolean,
	patch: (on: boolean) => StylePatch,
): Command =>
	editing({
		id,
		label,
		icon: icon(glyph),
		shortcut,
		lock: 'formatCells',
		checked: read,
		run: (ctx) => style(ctx, patch(!read(ctx))),
	});

export const BORDER_PRESETS: ReadonlyArray<readonly [BorderPreset, string, string]> = [
	['bottom', 'Bottom Border', 'borderBottom'],
	['top', 'Top Border', 'borderTop'],
	['left', 'Left Border', 'borderLeft'],
	['right', 'Right Border', 'borderRight'],
	['none', 'No Border', 'borderNone'],
	['all', 'All Borders', 'borderAll'],
	['outside', 'Outside Borders', 'borderOutside'],
	['thickOutside', 'Thick Outside Borders', 'borderThick'],
	['doubleBottom', 'Bottom Double Border', 'borderBottom'],
	['thickBottom', 'Thick Bottom Border', 'borderBottom'],
	['topAndBottom', 'Top and Bottom Border', 'borderOutside'],
	['inside', 'Inside Borders', 'borderInside'],
	['insideH', 'Inside Horizontal Borders', 'borderInside'],
	['insideV', 'Inside Vertical Borders', 'borderInside'],
];

let lastBorder: BorderPreset = 'bottom';

export function fontCommands(): Command[] {
	return [
		editing({
			id: 'home.font-name',
			label: 'Font',
			lock: 'formatCells',
			value: (ctx) => activeStyle(ctx)?.font.name ?? 'Calibri',
			run: (ctx, arg) => {
				if (typeof arg === 'string' && arg.trim())
					style(ctx, { font: { name: arg.trim(), scheme: UNSET } });
			},
		}),
		editing({
			id: 'home.font-size',
			label: 'Font Size',
			lock: 'formatCells',
			value: (ctx) => String(activeStyle(ctx)?.font.size ?? 11),
			run: (ctx, arg) => {
				const size = Number(arg);
				if (Number.isFinite(size) && size >= 1 && size <= 409)
					style(ctx, { font: { size: Math.round(size * 2) / 2 } });
			},
		}),
		...(['grow', 'shrink'] as const).map((dir) =>
			editing({
				id: `home.${dir}-font`,
				label: dir === 'grow' ? 'Increase Font Size' : 'Decrease Font Size',
				icon: icon(dir === 'grow' ? 'fontGrow' : 'fontShrink'),
				shortcut: dir === 'grow' ? 'Ctrl+Shift+>' : 'Ctrl+Shift+<',
				lock: 'formatCells',
				run: (ctx) =>
					style(ctx, {
						font: { size: stepFontSize(activeStyle(ctx)?.font.size ?? 11, dir === 'grow') },
					}),
			}),
		),
		toggle(
			'home.bold',
			'Bold',
			'bold',
			'Ctrl+B',
			(c) => !!activeStyle(c)?.font.bold,
			(on) => ({
				font: { bold: on || UNSET },
			}),
		),
		toggle(
			'home.italic',
			'Italic',
			'italic',
			'Ctrl+I',
			(c) => !!activeStyle(c)?.font.italic,
			(on) => ({
				font: { italic: on || UNSET },
			}),
		),
		toggle(
			'home.underline',
			'Underline',
			'underline',
			'Ctrl+U',
			(c) => activeStyle(c)?.font.underline === 'single',
			(on) => ({ font: { underline: on ? 'single' : UNSET } }),
		),
		toggle(
			'home.underline-double',
			'Double Underline',
			'underline',
			'',
			(c) => activeStyle(c)?.font.underline === 'double',
			(on) => ({ font: { underline: on ? 'double' : UNSET } }),
		),
		toggle(
			'home.strikethrough',
			'Strikethrough',
			'strike',
			'Ctrl+5',
			(c) => !!activeStyle(c)?.font.strike,
			(on) => ({
				font: { strike: on || UNSET },
			}),
		),
		editing({
			id: 'home.font-color',
			label: 'Font Color',
			icon: icon('fontColor'),
			lock: 'formatCells',
			run: (ctx, arg) => {
				const color = pickColor(arg);
				style(ctx, { font: { color: color ?? { theme: 1 } } });
			},
		}),
		editing({
			id: 'home.fill-color',
			label: 'Fill Color',
			icon: icon('fillColor'),
			lock: 'formatCells',
			run: (ctx, arg) => {
				const color = pickColor(arg);
				style(ctx, {
					fill: color
						? { type: 'pattern', pattern: 'solid', fgColor: color }
						: { type: 'pattern', pattern: 'none' },
				});
			},
		}),
		editing({
			id: 'home.borders',
			label: 'Borders',
			icon: icon('borderBottom'),
			lock: 'formatCells',
			run: (ctx, arg) => {
				const t = target(ctx);
				if (!t) return;
				const spec = parseBorderArg(arg);
				lastBorder = spec.preset;
				t.session.batch('Borders', () => {
					for (const range of t.ranges)
						t.session.setBorders(t.sheet, range, spec.preset, spec.edge);
				});
			},
		}),
		editing({
			id: 'home.font-settings',
			label: 'Font Settings',
			icon: icon('formatCells'),
			lock: 'formatCells',
			run: (ctx) => void ctx.dialogs.open('format-cells', { tab: 'font' }),
		}),
		editing({
			id: 'home.borders-more',
			label: 'More Borders...',
			icon: icon('borderMore'),
			lock: 'formatCells',
			run: (ctx) => void ctx.dialogs.open('format-cells', { tab: 'border' }),
		}),
	];
}

/** `home.borders` accepts a preset name or `{ preset, style, color }`; nothing re-applies the last one. */
export function parseBorderArg(arg: unknown): { preset: BorderPreset; edge?: BorderEdge } {
	if (typeof arg === 'string' && BORDER_PRESETS.some(([p]) => p === arg))
		return { preset: arg as BorderPreset };
	if (arg && typeof arg === 'object' && 'preset' in arg) {
		const spec = arg as { preset: BorderPreset; style?: BorderEdge['style']; color?: Color };
		const edge: BorderEdge = { style: spec.style ?? 'thin' };
		if (spec.color) edge.color = spec.color;
		return { preset: spec.preset, edge };
	}
	return { preset: lastBorder };
}
