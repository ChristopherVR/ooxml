/**
 * Excel's colour picker: Automatic (or No Fill), Theme Colors (ten columns, each a theme slot over
 * five tints taken from the workbook theme), Standard Colors and More Colors (the browser's colour
 * input). The choice is a SpreadsheetML `Color` (`{ theme, tint }` or `{ rgb }`), or undefined.
 */
import {
	DEFAULT_THEME,
	resolveColor,
	type Color,
	type ThemePalette,
} from '@christophervr/xlsx-core';
import { arrowNavigation, closeRibbonPopover, mountPopover } from './popover';

/** Theme slots in Excel's column order with their role names. */
const THEME_COLUMNS: ReadonlyArray<readonly [number, string]> = [
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

/** Excel's five tint rows under each theme colour (lt1, dk1, lt2 and the rest differ). */
export function themeTints(slot: number): number[] {
	if (slot === 0) return [-0.05, -0.15, -0.25, -0.35, -0.5];
	if (slot === 1) return [0.5, 0.35, 0.25, 0.15, 0.05];
	if (slot === 2) return [-0.1, -0.25, -0.5, -0.75, -0.9];
	return [0.8, 0.6, 0.4, -0.25, -0.5];
}

export const STANDARD_COLORS: ReadonlyArray<readonly [string, string]> = [
	['C00000', 'Dark Red'],
	['FF0000', 'Red'],
	['FFC000', 'Orange'],
	['FFFF00', 'Yellow'],
	['92D050', 'Light Green'],
	['00B050', 'Green'],
	['00B0F0', 'Light Blue'],
	['0070C0', 'Blue'],
	['002060', 'Dark Blue'],
	['7030A0', 'Purple'],
];

export interface ColorGridOptions {
	/** First row: 'Automatic' (font colour) or 'No Fill' (fill colour). */
	automaticLabel?: string;
	theme?: ThemePalette;
	t(key: string, vars?: Record<string, string | number>): string;
}

const tintName = (tint: number, t: ColorGridOptions['t']) =>
	tint > 0
		? t('Lighter {percent}%', { percent: Math.round(tint * 100) })
		: t('Darker {percent}%', { percent: Math.round(-tint * 100) });

/** The CSS colour a `Color` shows in the workbook theme (for swatches and the split bar). */
export function cssColor(color: Color | undefined, theme: ThemePalette = DEFAULT_THEME): string {
	return resolveColor(color, theme) ?? 'transparent';
}

export function openColorGrid(
	anchor: HTMLElement,
	choose: (color: Color | undefined) => void,
	options: ColorGridOptions,
): void {
	const { t } = options;
	const theme = options.theme ?? DEFAULT_THEME;
	const doc = anchor.ownerDocument;
	const pop = doc.createElement('div');
	pop.className = 'ribbon-popover color-grid';
	pop.setAttribute('role', 'menu');
	const pick = (color: Color | undefined) => {
		closeRibbonPopover();
		choose(color);
	};
	const swatch = (color: Color, name: string) => {
		const item = doc.createElement('button');
		item.type = 'button';
		item.setAttribute('role', 'menuitem');
		item.className = 'swatch';
		item.style.setProperty('--swatch', cssColor(color, theme));
		item.setAttribute('aria-label', name);
		item.title = name;
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => pick(color));
		return item;
	};
	const title = (text: string) => {
		const el = doc.createElement('div');
		el.className = 'color-grid-title';
		el.textContent = t(text);
		return el;
	};
	const cells = (items: HTMLElement[]) => {
		const el = doc.createElement('div');
		el.className = 'color-grid-cells';
		el.append(...items);
		return el;
	};
	const command = (text: string, run: () => void) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.setAttribute('role', 'menuitem');
		button.className = 'color-grid-command';
		button.textContent = text;
		button.addEventListener('mousedown', (event) => event.preventDefault());
		button.addEventListener('click', run);
		return button;
	};
	if (options.automaticLabel) pop.append(command(t(options.automaticLabel), () => pick(undefined)));
	const top = THEME_COLUMNS.map(([slot, name]) => swatch({ theme: slot }, t(name)));
	const rows = [0, 1, 2, 3, 4].flatMap((row) =>
		THEME_COLUMNS.map(([slot, name]) => {
			const tint = themeTints(slot)[row] ?? 0;
			return swatch({ theme: slot, tint }, `${t(name)}, ${tintName(tint, t)}`);
		}),
	);
	pop.append(
		title('Theme Colors'),
		cells([...top, ...rows]),
		title('Standard Colors'),
		cells(STANDARD_COLORS.map(([rgb, name]) => swatch({ rgb }, t(name)))),
	);
	const input = doc.createElement('input');
	input.type = 'color';
	input.className = 'color-grid-input';
	input.tabIndex = -1;
	input.addEventListener('change', () => pick({ rgb: input.value.slice(1).toUpperCase() }));
	pop.append(
		command(t('More Colors...'), () => input.click()),
		input,
	);
	arrowNavigation(pop, '[role="menuitem"]');
	if (!mountPopover(anchor, pop)) return;
	pop.querySelector<HTMLElement>('[role="menuitem"]')?.focus();
}
