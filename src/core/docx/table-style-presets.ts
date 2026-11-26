import type { TableBorderSide, TableStyleCatalog, TableStyleDefinition } from './table-model';
import { DEFAULT_TABLE_BORDERS } from './table-defaults';
import { eighthPoints } from './units';

const line = (color: string): TableBorderSide => ({
	style: 'single',
	sizeEighthPoints: eighthPoints(4),
	color,
});

// Independently defined previews and OOXML formatting for common Word gallery entries.
// Names follow the public Word gallery and WdBuiltinStyle.
export const TABLE_STYLE_PRESETS: readonly TableStyleDefinition[] = [
	{ id: 'TableNormal', name: 'Normal Table', conditional: {} },
	{ id: 'TableGrid', name: 'Table Grid', borders: DEFAULT_TABLE_BORDERS, conditional: {} },
	{
		id: 'LightShading',
		name: 'Light Shading',
		borders: { top: line('#000000'), bottom: line('#000000') },
		conditional: {
			firstRow: { borders: { bottom: line('#000000') }, run: { bold: true } },
			lastRow: { borders: { top: line('#000000') }, run: { bold: true } },
			firstCol: { run: { bold: true } },
			lastCol: { run: { bold: true } },
			band1Horz: { shadingFill: '#E7E6E6' },
			band1Vert: { shadingFill: '#E7E6E6' },
		},
	},
	{
		id: 'LightList',
		name: 'Light List',
		borders: { bottom: line('#000000'), insideH: line('#000000') },
		conditional: {
			firstRow: { shadingFill: '#000000', run: { bold: true, color: '#FFFFFF' } },
			lastRow: { borders: { top: line('#000000') }, run: { bold: true } },
			firstCol: { run: { bold: true } },
			lastCol: { run: { bold: true } },
		},
	},
	{
		id: 'LightGrid',
		name: 'Light Grid',
		borders: DEFAULT_TABLE_BORDERS,
		conditional: {
			firstRow: { shadingFill: '#D9D9D9', run: { bold: true } },
			lastRow: { shadingFill: '#D9D9D9', run: { bold: true } },
			firstCol: { run: { bold: true } },
			lastCol: { run: { bold: true } },
			band1Horz: { shadingFill: '#F2F2F2' },
			band1Vert: { shadingFill: '#F2F2F2' },
		},
	},
	{
		id: 'GridTable4-Accent1',
		name: 'Grid Table 4 - Accent 1',
		borders: {
			top: line('#4472C4'),
			bottom: line('#4472C4'),
			left: line('#4472C4'),
			right: line('#4472C4'),
			insideH: line('#8EAADB'),
			insideV: line('#8EAADB'),
		},
		conditional: {
			firstRow: { shadingFill: '#4472C4', run: { bold: true, color: '#FFFFFF' } },
			lastRow: { shadingFill: '#B4C6E7', run: { bold: true } },
			firstCol: { run: { bold: true } },
			lastCol: { run: { bold: true } },
			band1Horz: { shadingFill: '#D9E2F3' },
			band1Vert: { shadingFill: '#D9E2F3' },
		},
	},
];

/** A built-in selected in the gallery becomes a real, self-contained document style. */
export function installTableStylePreset(
	catalog: TableStyleCatalog | undefined,
	preset: TableStyleDefinition,
): TableStyleCatalog {
	if (catalog && Object.hasOwn(catalog.styles, preset.id)) return catalog;
	return {
		warnings: catalog?.warnings ?? [],
		styles: { ...catalog?.styles, [preset.id]: structuredClone(preset) },
	};
}
