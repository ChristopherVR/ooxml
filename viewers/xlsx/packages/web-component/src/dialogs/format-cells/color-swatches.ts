// A keyboard-accessible colour swatch grid for the Format Cells tabs and the Tab Color dialog:
// an Automatic / No Color button, the theme colours with their tint rows and the standard colours.
import type { Color } from '@christophervr/xlsx-core';
import { resolveColor } from '@christophervr/xlsx-core';
import type { EditorContext } from '../../context.js';
import { STANDARD_COLORS, themeTints } from '../../ribbon/color-grid.js';
import { el } from '../fields.js';

const THEME_NAMES: ReadonlyArray<readonly [number, string]> = [
	[0, 'Background 1'],
	[1, 'Text 1'],
	[2, 'Background 2'],
	[3, 'Text 2'],
	[4, 'Accent 1'],
	[5, 'Accent 2'],
	[6, 'Accent 3'],
	[7, 'Accent 4'],
	[8, 'Accent 5'],
	[9, 'Accent 6'],
];

export const colorKey = (color: Color | undefined): string =>
	color
		? JSON.stringify([
				color.rgb?.slice(-6).toUpperCase(),
				color.theme,
				color.tint,
				color.indexed,
				color.auto,
			])
		: 'none';

export interface SwatchGrid {
	element: HTMLDivElement;
	value(): Color | undefined;
	set(color: Color | undefined): void;
}

/** `noneLabel` names the "no colour" button (`Automatic` or `No Color`). */
export function swatchGrid(
	ctx: EditorContext,
	label: string,
	noneLabel: string,
	initial: Color | undefined,
	onChange?: (color: Color | undefined) => void,
): SwatchGrid {
	const theme = ctx.workbook()?.theme ?? { colors: [], majorFont: '', minorFont: '' };
	const wrapper = el(ctx, 'div', 'xve-swatch-picker');
	wrapper.setAttribute('role', 'group');
	wrapper.setAttribute('aria-label', ctx.t(label));
	let current = initial;
	const buttons: { button: HTMLButtonElement; color: Color | undefined }[] = [];
	const paint = (): void => {
		const key = colorKey(current);
		for (const { button, color } of buttons)
			button.setAttribute('aria-pressed', String(colorKey(color) === key));
	};
	const choose = (color: Color | undefined): void => {
		current = color;
		paint();
		onChange?.(color);
	};
	const none = el(ctx, 'button', 'xve-btn xve-swatch-none');
	none.type = 'button';
	none.textContent = ctx.t(noneLabel);
	none.addEventListener('click', () => choose(undefined));
	buttons.push({ button: none, color: undefined });
	const swatch = (color: Color, name: string): HTMLButtonElement => {
		const b = el(ctx, 'button', 'xve-swatch');
		b.type = 'button';
		b.style.background = resolveColor(color, theme) ?? 'transparent';
		b.setAttribute('aria-label', name);
		b.title = name;
		b.addEventListener('click', () => choose(color));
		buttons.push({ button: b, color });
		return b;
	};
	const tintName = (tint: number): string =>
		tint > 0
			? ctx.t('Lighter {percent}%', { percent: Math.round(tint * 100) })
			: ctx.t('Darker {percent}%', { percent: Math.round(-tint * 100) });
	const themeGrid = el(ctx, 'div', 'xve-swatches');
	for (const [slot, name] of THEME_NAMES) themeGrid.append(swatch({ theme: slot }, ctx.t(name)));
	for (let row = 0; row < 5; row++)
		for (const [slot, name] of THEME_NAMES) {
			const tint = themeTints(slot)[row] ?? 0;
			themeGrid.append(swatch({ theme: slot, tint }, `${ctx.t(name)}, ${tintName(tint)}`));
		}
	const standard = el(ctx, 'div', 'xve-swatches');
	for (const [rgb, name] of STANDARD_COLORS)
		standard.append(swatch({ rgb: `FF${rgb}` }, ctx.t(name)));
	const cap = (key: string): HTMLDivElement => {
		const d = el(ctx, 'div', 'xve-note');
		d.textContent = ctx.t(key);
		return d;
	};
	wrapper.append(none, cap('Theme Colors'), themeGrid, cap('Standard Colors'), standard);
	// Arrow keys move between swatches (ten per row).
	wrapper.addEventListener('keydown', (event) => {
		const list = buttons.map((b) => b.button);
		const index = list.indexOf(event.target as HTMLButtonElement);
		if (index < 0) return;
		const delta = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: 10, ArrowUp: -10 }[event.key];
		if (!delta) return;
		event.preventDefault();
		const next =
			index === 0 ? (delta > 0 ? 1 : 0) : Math.max(0, Math.min(list.length - 1, index + delta));
		list[next]?.focus();
	});
	paint();
	return {
		element: wrapper,
		value: () => current,
		set(color) {
			current = color;
			paint();
		},
	};
}
