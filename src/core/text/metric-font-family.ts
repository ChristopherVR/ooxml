import { buildFontFamilyString } from './font-substitution';
export const METRIC_FONT_SUBSTITUTES: Record<string, { fallbacks: string[]; generic: string }> = {
	calibri: { fallbacks: ['Carlito'], generic: 'sans-serif' },
	'calibri light': { fallbacks: ['Carlito'], generic: 'sans-serif' },
	cambria: { fallbacks: ['Caladea'], generic: 'serif' },
	arial: { fallbacks: ['Arimo', 'Liberation Sans'], generic: 'sans-serif' },
	helvetica: { fallbacks: ['Arimo', 'Liberation Sans'], generic: 'sans-serif' },
	'times new roman': { fallbacks: ['Tinos', 'Liberation Serif'], generic: 'serif' },
	times: { fallbacks: ['Tinos', 'Liberation Serif'], generic: 'serif' },
	'courier new': { fallbacks: ['Cousine', 'Liberation Mono'], generic: 'monospace' },
	courier: { fallbacks: ['Cousine', 'Liberation Mono'], generic: 'monospace' },
	georgia: { fallbacks: ['Gelasio'], generic: 'serif' },
	consolas: { fallbacks: ['Inconsolata'], generic: 'monospace' },
};

/** Metric-compatible policy used where stable pagination is required. */
export function metricFontFamily(name: string): string {
	const key = name.trim();
	const known = Object.hasOwn(METRIC_FONT_SUBSTITUTES, key.toLowerCase())
		? METRIC_FONT_SUBSTITUTES[key.toLowerCase()]
		: undefined;
	return buildFontFamilyString(key, [...(known?.fallbacks ?? []), known?.generic ?? 'sans-serif']);
}
