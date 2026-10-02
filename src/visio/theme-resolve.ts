/**
 * MS-VSDX ColorSchemeIndex (2.4.4.58), QuickStyleFillColor (2.4.4.270),
 * QuickStyleFillMatrix (2.4.4.271), and QuickStyleType (2.4.4.277):
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/1e7e9b7e-d116-41c0-9c53-5e57c26042a4
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/723529ce-4ac3-416d-b7be-362aa28be341
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/25689058-b1e7-4d3c-a833-0a4c7180f5f2
 * https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/550dd5fb-2388-41c9-867d-3a8af03c94b8
 */
import type { Cells, Report, Sheet } from './sheet.js';
import type { VisioTheme } from './theme.js';
import { colorChoice, drawingColor, integer, themeChild, THEME_NS } from './theme-color.js';

export interface ThemeResources {
	themes?: VisioTheme[];
	pageCells?: Cells;
	/** Verified ID 0 / NameU No Style, for explicitly selected root formatting. */
	rootSheet?: Sheet;
}
const cellNumber = (cells: Cells | undefined, name: string) => integer(cells?.get(name)?.value);
/** Only saved zero, or saved page inheritance to zero, explicitly selects root formats. */
export function rootThemeSelected(cells: Cells, resources: ThemeResources): boolean {
	const local = cellNumber(cells, 'ColorSchemeIndex');
	return (
		local === 0 || (local === 65534 && cellNumber(resources.pageCells, 'ColorSchemeIndex') === 0)
	);
}

function inheritedIndex(
	cells: Cells,
	page: Cells | undefined,
	name: string,
	fallback: number,
): number | undefined {
	const local = cellNumber(cells, name);
	if (cells.has(name) && local === undefined) return undefined;
	if (local !== undefined && local !== 65534) return local;
	const value = cellNumber(page, name);
	return page?.has(name) && value === undefined ? undefined : (value ?? fallback);
}
function themeFor(
	themes: VisioTheme[],
	id: number | undefined,
	key: 'colorId' | 'effectId' | 'connectorId',
): VisioTheme | undefined {
	if (id === undefined || id === 0 || id >= 65534) return undefined;
	const matching = themes.filter((theme) => theme[key] === id);
	return matching.length === 1 ? matching[0] : undefined;
}
function quickColor(
	theme: VisioTheme,
	index: number | undefined,
	variant: number | undefined,
	report: Report,
): string | undefined {
	if (index === undefined) return undefined;
	const name = [
		'dk1',
		'lt1',
		'accent1',
		'accent2',
		'accent3',
		'accent4',
		'accent5',
		'accent6',
		'bkgnd',
	][index];
	if (name) return drawingColor(theme.colors.get(name), theme.colors);
	// Observed in Office-generated VSDX. This extension is not enumerated by MS-VSDX
	// 2.4.4.270, so preserve a diagnostic rather than claiming full theme parity.
	if (index >= 200 && index <= 206) {
		report(
			'theme-color-extension',
			'Quick Style color 200-206 is interpreted as a theme variant color.',
		);
		index -= 100;
	}
	if (index < 100 || index > 106 || variant === undefined || variant > 3) return undefined;
	return drawingColor(theme.variants[variant]?.get(String(index - 99)), theme.colors);
}
export type ThemeCategory = 'Fill' | 'Line' | 'Font' | 'Effects';
/** Select saved format indices independently of colors, except explicit root-format selection. */
export function themeFormat(
	cells: Cells,
	category: ThemeCategory,
	resources: ThemeResources,
): Element | undefined {
	// Root selection is authoritative even when its saved cache is absent or unusable.
	if (rootThemeSelected(cells, resources) || !resources.themes?.length) return undefined;
	const type = cells.has('QuickStyleType') ? cellNumber(cells, 'QuickStyleType') : 0;
	if (type === undefined || type > 3) return undefined;
	const oneD = cellNumber(cells, 'OneD') === 1 || (cells.has('BeginX') && cells.has('EndX'));
	const connector = type === 3 || (type === 0 && oneD);
	const scheme = connector ? 'ConnectorSchemeIndex' : 'EffectSchemeIndex';
	const effect = themeFor(
		resources.themes,
		inheritedIndex(cells, resources.pageCells, scheme, 0),
		connector ? 'connectorId' : 'effectId',
	);
	if (!effect) return undefined;
	let matrix = cellNumber(cells, `QuickStyle${category}Matrix`);
	if (matrix !== undefined && matrix >= 100 && matrix <= 103) {
		const variant = inheritedIndex(cells, resources.pageCells, 'VariationStyleIndex', 0);
		const style =
			variant === undefined ? undefined : effect.variationStyles[variant]?.[matrix - 100];
		matrix = integer(
			style?.getAttribute(`${category === 'Effects' ? 'effect' : category.toLowerCase()}Idx`),
		);
	}
	if (matrix === undefined || matrix < 1 || matrix > 6) return undefined;
	const list =
		category === 'Effects'
			? connector
				? effect.connectorEffects
				: effect.effects
			: category === 'Fill'
				? connector
					? effect.connectorFills
					: effect.fills
				: category === 'Line'
					? connector
						? effect.connectorLines
						: effect.lines
					: connector
						? effect.connectorFonts
						: effect.fonts;
	return list[matrix - 1];
}
export function themePaintContext(
	cells: Cells,
	category: ThemeCategory,
	resources: ThemeResources,
	report: Report,
):
	| {
			selected: Element;
			colors: ReadonlyMap<string, Element>;
			base: string | undefined;
	  }
	| undefined {
	if (!resources.themes?.length) return undefined;
	const variation = cells.has('QuickStyleVariation') ? cellNumber(cells, 'QuickStyleVariation') : 0;
	const visibilityBit = category === 'Fill' ? 8 : category === 'Line' ? 4 : 2;
	if (variation === undefined || variation > 15 || variation & visibilityBit) return undefined;
	const colorId = inheritedIndex(cells, resources.pageCells, 'ColorSchemeIndex', 0);
	const colorTheme = themeFor(resources.themes, colorId, 'colorId');
	const selected = themeFormat(cells, category, resources);
	if (!colorTheme || !selected) return undefined;
	const variant = inheritedIndex(cells, resources.pageCells, 'VariationColorIndex', 0);
	return {
		selected,
		colors: colorTheme.colors,
		base: quickColor(colorTheme, cellNumber(cells, `QuickStyle${category}Color`), variant, report),
	};
}
/** Resolve only saved theme selectors, never ShapeSheet formula text. */
export function themeColor(
	cells: Cells,
	name: string,
	resources: ThemeResources,
	report: Report,
	gradientRendered = false,
): string | undefined {
	const category =
		name === 'FillForegnd'
			? 'Fill'
			: name === 'LineColor'
				? 'Line'
				: name === 'Color'
					? 'Font'
					: undefined;
	if (!category) return undefined;
	const context = themePaintContext(cells, category, resources, report);
	if (!context) return undefined;
	const { selected, colors, base } = context;
	if (category === 'Font')
		return drawingColor(colorChoice(themeChild(selected, 'color', THEME_NS)), colors, base);
	const fill =
		category === 'Fill'
			? selected
			: (themeChild(selected, 'solidFill') ??
				themeChild(selected, 'noFill') ??
				themeChild(selected, 'gradFill'));
	if (fill?.localName === 'solidFill') return drawingColor(colorChoice(fill), colors, base);
	if (fill?.localName === 'gradFill' && base) {
		if (!gradientRendered)
			report(
				'unsupported-theme-gradient',
				'A theme gradient is approximated by its untransformed Quick Style color.',
			);
		return base;
	}
	return undefined;
}
