import type { XlsxTheme } from './types';

/** Every token key; each becomes the kebab-case `--xve-*` custom property. */
export const THEME_KEYS = [
	'background',
	'foreground',
	'card',
	'cardForeground',
	'popover',
	'popoverForeground',
	'primary',
	'primaryForeground',
	'secondary',
	'secondaryForeground',
	'muted',
	'mutedForeground',
	'accent',
	'accentForeground',
	'destructive',
	'destructiveForeground',
	'border',
	'input',
	'ring',
	'radius',
	'gridLine',
	'headerBg',
	'headerFg',
	'headerActiveBg',
	'selection',
	'selectionBorder',
	'sheetBg',
] as const satisfies readonly (keyof XlsxTheme)[];

const kebab = (key: string) => key.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`);

/** Converts a (partial) theme to `--xve-*` custom properties; unknown or empty values are skipped. */
export function themeToCssVars(theme: Partial<XlsxTheme> | undefined): Record<string, string> {
	const vars: Record<string, string> = {};
	if (!theme) return vars;
	for (const key of THEME_KEYS) {
		const value = theme[key];
		if (typeof value === 'string' && value.trim()) vars[`--xve-${kebab(key)}`] = value.trim();
	}
	return vars;
}
