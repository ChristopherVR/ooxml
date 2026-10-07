import type { ChartStyleEntry, ChartStylePart } from './style-definition';

const TEXT_KEYS = ['fontSize', 'bold', 'italic', 'typeface', 'textColor', 'fontRef'] as const;
const TEXT_PARTS: readonly ChartStylePart[] = [
	'title',
	'axisTitle',
	'categoryAxis',
	'valueAxis',
	'legend',
	'dataLabel',
];

/** A chart-space text body supplies inherited text, while element formatting remains direct. */
export function inheritChartText(
	part: ChartStylePart,
	entry: ChartStyleEntry,
	builtIn: ChartStyleEntry | undefined,
	root: ChartStyleEntry | undefined,
): ChartStyleEntry {
	if (!root || !TEXT_PARTS.includes(part)) return entry;
	const out = { ...entry };
	for (const key of TEXT_KEYS) {
		delete out[key];
		const value = root[key] ?? builtIn?.[key];
		if (value !== undefined) Object.assign(out, { [key]: value });
	}
	// Native inherited title size is scaled; explicit title sizes are applied afterward.
	if (part === 'title' && root.fontSize !== undefined) out.fontSize = root.fontSize * 1.2;
	if (root.textEffectsXml !== undefined) out.textEffectsXml = root.textEffectsXml;
	return out;
}
