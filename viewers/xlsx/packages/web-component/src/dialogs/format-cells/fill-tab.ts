// Format Cells > Fill: background colour, pattern colour and pattern style with a sample. Cell
// fills store a solid colour as fgColor; conditional-format (dxf) fills store it as bgColor.
import type { Color, Fill, PatternType } from '@christophervr/xlsx-core';
import { resolveColor } from '@christophervr/xlsx-core';
import { el, field, fieldset, row, select } from '../fields.js';
import { swatchGrid } from './color-swatches.js';
import type { FormatTab, TabInit } from './types.js';

export const PATTERNS: ReadonlyArray<readonly [PatternType, string]> = [
	['none', 'None'],
	['solid', 'Solid'],
	['darkGray', '75% Gray'],
	['mediumGray', '50% Gray'],
	['lightGray', '25% Gray'],
	['gray125', '12.5% Gray'],
	['gray0625', '6.25% Gray'],
	['darkHorizontal', 'Horizontal Stripe'],
	['darkVertical', 'Vertical Stripe'],
	['darkDown', 'Reverse Diagonal Stripe'],
	['darkUp', 'Diagonal Stripe'],
	['darkGrid', 'Diagonal Crosshatch'],
	['darkTrellis', 'Thick Diagonal Crosshatch'],
	['lightHorizontal', 'Thin Horizontal Stripe'],
	['lightVertical', 'Thin Vertical Stripe'],
	['lightDown', 'Thin Reverse Diagonal Stripe'],
	['lightUp', 'Thin Diagonal Stripe'],
	['lightGrid', 'Thin Horizontal Crosshatch'],
	['lightTrellis', 'Thin Diagonal Crosshatch'],
];

/** The dialog's three fields for a fill. */
export function readFill(
	fill: Fill,
	dxf: boolean,
): { background?: Color; patternColor?: Color; pattern: PatternType } {
	if (fill.type !== 'pattern' || fill.pattern === 'none') return { pattern: 'none' };
	if (fill.pattern === 'solid') {
		const background = dxf ? (fill.bgColor ?? fill.fgColor) : (fill.fgColor ?? fill.bgColor);
		return background ? { background, pattern: 'none' } : { pattern: 'none' };
	}
	const out: { background?: Color; patternColor?: Color; pattern: PatternType } = {
		pattern: fill.pattern,
	};
	if (fill.bgColor) out.background = fill.bgColor;
	if (fill.fgColor) out.patternColor = fill.fgColor;
	return out;
}

export function buildFill(
	background: Color | undefined,
	patternColor: Color | undefined,
	pattern: PatternType,
	dxf: boolean,
): Fill {
	if (pattern === 'none' || pattern === 'solid') {
		const color = background ?? (pattern === 'solid' ? patternColor : undefined);
		if (!color) return { type: 'pattern', pattern: 'none' };
		return dxf
			? { type: 'pattern', pattern: 'solid', bgColor: color }
			: { type: 'pattern', pattern: 'solid', fgColor: color };
	}
	const fill: Fill = { type: 'pattern', pattern, fgColor: patternColor ?? { indexed: 64 } };
	if (background) fill.bgColor = background;
	return fill;
}

export function fillTab(init: TabInit): FormatTab {
	const { ctx } = init;
	const theme = ctx.workbook()?.theme ?? { colors: [], majorFont: '', minorFont: '' };
	const start = readFill(init.style.fill, init.dxf);
	let dirty = false;
	const mark = (): void => {
		dirty = true;
		paint();
	};
	const background = swatchGrid(ctx, 'Background Color:', 'No Color', start.background, mark);
	const patternColor = swatchGrid(ctx, 'Pattern Color:', 'Automatic', start.patternColor, mark);
	const pattern = select(ctx, PATTERNS, start.pattern);
	pattern.addEventListener('change', mark);
	const sample = el(ctx, 'div', 'xve-sample');
	sample.dataset.sample = '';
	const paint = (): void => {
		const bg = resolveColor(background.value(), theme) ?? '#ffffff';
		const fg = resolveColor(patternColor.value(), theme) ?? '#000000';
		const p = pattern.value as PatternType;
		sample.style.background =
			p === 'none' || p === 'solid'
				? bg
				: `repeating-linear-gradient(45deg, ${fg} 0 2px, ${bg} 2px 5px)`;
	};
	paint();
	const current = (): Fill =>
		buildFill(background.value(), patternColor.value(), pattern.value as PatternType, init.dxf);
	const panel = el(ctx, 'div');
	panel.append(
		row(
			ctx,
			fieldset(ctx, 'Background Color:', background.element),
			fieldset(ctx, 'Pattern', field(ctx, 'Pattern Style:', pattern), patternColor.element),
		),
		fieldset(ctx, 'Sample', sample),
	);
	return {
		id: 'fill',
		label: 'Fill',
		panel,
		dirty: () => dirty,
		patch: () => (dirty ? { fill: current() } : {}),
		toDxf: (dxf) => {
			const fill = current();
			if (fill.type === 'pattern' && fill.pattern === 'none') delete dxf.fill;
			else dxf.fill = fill;
		},
		focus: () => background.element.querySelector<HTMLButtonElement>('button')?.focus(),
	};
}
