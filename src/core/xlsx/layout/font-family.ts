/** Workbook font stacks, shared by cell and chart painters. */
const FALLBACKS: Record<string, string> = {
	calibri: '"Calibri", "Carlito", "Segoe UI", Arial, sans-serif',
	'calibri light': '"Calibri Light", "Calibri", "Carlito", "Segoe UI", Arial, sans-serif',
	cambria: '"Cambria", "Caladea", Georgia, serif',
	arial: 'Arial, "Liberation Sans", Helvetica, sans-serif',
	'times new roman': '"Times New Roman", "Liberation Serif", Times, serif',
	'courier new': '"Courier New", "Liberation Mono", monospace',
	aptos: '"Aptos", "Calibri", "Carlito", "Segoe UI", sans-serif',
};

export function cssFontFamily(name: string): string {
	const known = FALLBACKS[name.toLowerCase()];
	if (known) return known;
	const quoted = `"${name.replace(/["\\]/g, '')}"`;
	return `${quoted}, "Calibri", "Carlito", "Segoe UI", sans-serif`;
}
