/**
 * Workbook font stacks, shared by cell and chart painters. Each Office font lists the verified
 * metric-compatible clone the webfont fallback downloads (`webfonts/metric-clones.ts`) right after
 * it, so a downloaded clone measures and paints in place of a missing original.
 */
const FALLBACKS: Record<string, string> = {
	calibri: '"Calibri", "Carlito", "Segoe UI", Arial, sans-serif',
	'calibri light': '"Calibri Light", "Calibri", "Carlito", "Segoe UI", Arial, sans-serif',
	cambria: '"Cambria", "Caladea", Georgia, serif',
	arial: 'Arial, "Arimo", "Liberation Sans", Helvetica, sans-serif',
	'times new roman': '"Times New Roman", "Tinos", "Liberation Serif", Times, serif',
	'courier new': '"Courier New", "Cousine", "Liberation Mono", monospace',
	georgia: '"Georgia", "Gelasio", serif',
	aptos: '"Aptos", "Calibri", "Carlito", "Segoe UI", sans-serif',
	'aptos narrow': '"Aptos Narrow", "Aptos", "Calibri", "Carlito", "Segoe UI", sans-serif',
	'aptos display': '"Aptos Display", "Aptos", "Calibri", "Carlito", "Segoe UI", sans-serif',
};

export function cssFontFamily(name: string): string {
	const known = FALLBACKS[name.toLowerCase()];
	if (known) return known;
	const quoted = `"${name.replace(/["\\]/g, '')}"`;
	return `${quoted}, "Calibri", "Carlito", "Segoe UI", sans-serif`;
}
