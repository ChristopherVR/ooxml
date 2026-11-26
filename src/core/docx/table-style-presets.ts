import type { TableBorderSide, TableStyleCatalog, TableStyleDefinition } from './table-model';
import { DEFAULT_TABLE_BORDERS } from './table-defaults';
import { applyThemeTint } from './theme-color';
import type { ThemeColorToken } from './theme-model';
import { eighthPoints } from './units';

const line = (color: string, themeColor: ThemeColorToken = 'text1'): TableBorderSide => ({
	style: 'single',
	sizeEighthPoints: eighthPoints(4),
	color,
	themeColor,
});

const fill = (token: ThemeColorToken, color: string, tint?: number) => ({
	shadingFill: tint === undefined ? color : applyThemeTint(color, tint),
	shadingThemeFill: { token, ...(tint === undefined ? {} : { tint }) },
});
const black = line('#000000');
const lightText = { bold: true, color: '#FFFFFF', colorTheme: { token: 'background1' } } as const;
const plainBand = (tint: number) => fill('text1', '#000000', tint);

function accentStyle(family: 'Grid' | 'List', accent: 1 | 2, color: string): TableStyleDefinition {
	const themeColor = accent === 1 ? 'accent1' : 'accent2';
	const edge = line(color, themeColor);
	return {
		id: `${family}Table4-Accent${accent}`,
		name: `${family} Table 4 - Accent ${accent}`,
		borders: {
			top: edge,
			bottom: edge,
			insideH: edge,
			...(family === 'Grid' && { left: edge, right: edge, insideV: edge }),
		},
		conditional: {
			firstRow: { ...fill(themeColor, color), run: lightText },
			lastRow: { ...fill(themeColor, color, 0.4), run: { bold: true } },
			firstCol: { run: { bold: true } },
			lastCol: { run: { bold: true } },
			band1Horz: fill(themeColor, color, 0.2),
			band1Vert: fill(themeColor, color, 0.2),
		},
	};
}

// Independently defined previews and OOXML formatting for common Word gallery entries.
// Names follow the public Word gallery and WdBuiltinStyle.
export const TABLE_STYLE_PRESETS: readonly TableStyleDefinition[] = [
	{ id: 'TableNormal', name: 'Normal Table', conditional: {} },
	{
		id: 'PlainTable1',
		name: 'Plain Table 1',
		conditional: { firstRow: { borders: { bottom: black }, run: { bold: true } } },
	},
	{
		id: 'PlainTable2',
		name: 'Plain Table 2',
		borders: { top: black, bottom: black },
		conditional: {
			firstRow: { ...plainBand(0.15), borders: { bottom: black }, run: { bold: true } },
		},
	},
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
			band1Horz: plainBand(0.1),
			band1Vert: plainBand(0.1),
		},
	},
	{
		id: 'LightList',
		name: 'Light List',
		borders: { bottom: line('#000000'), insideH: line('#000000') },
		conditional: {
			firstRow: { ...fill('text1', '#000000'), run: lightText },
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
			firstRow: { ...plainBand(0.15), run: { bold: true } },
			lastRow: { ...plainBand(0.15), run: { bold: true } },
			firstCol: { run: { bold: true } },
			lastCol: { run: { bold: true } },
			band1Horz: plainBand(0.05),
			band1Vert: plainBand(0.05),
		},
	},
	accentStyle('Grid', 1, '#4472C4'),
	accentStyle('Grid', 2, '#ED7D31'),
	accentStyle('List', 1, '#4472C4'),
	accentStyle('List', 2, '#ED7D31'),
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
