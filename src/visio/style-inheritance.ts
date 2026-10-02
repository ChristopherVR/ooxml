import { VisioPackageError } from './package.js';
import { emptySheet, mergeSheets, number, type Report, type Sheet } from './sheet.js';

export interface StyleRecord {
	sheet: Sheet;
	attributes: Map<string, string>;
}
export interface StyleContext {
	styles: Map<string, StyleRecord>;
	styleCache: Map<string, Sheet>;
	report: Report;
	checkTime: () => void;
}
export function styleSheet(
	id: string,
	category: string,
	context: StyleContext,
	visited = new Set<string>(),
): Sheet {
	context.checkTime();
	const key = `${category}:${id}`;
	const cached = context.styleCache.get(key);
	if (cached) return cached;
	if (visited.has(key)) {
		context.report('style-inheritance-cycle', `Style ${id} has cyclic inheritance.`);
		return emptySheet();
	}
	const style = context.styles.get(id);
	if (!style) {
		context.report('missing-style', `Style ${id} could not be resolved.`);
		return emptySheet();
	}
	const enabledCell =
		category === 'LineStyle'
			? 'EnableLineProps'
			: category === 'FillStyle'
				? 'EnableFillProps'
				: 'EnableTextProps';
	if (number(style.sheet.cells, enabledCell, 1, context.report) === 0) return emptySheet();
	if (visited.size > 64)
		throw new VisioPackageError('STYLE_DEPTH_LIMIT', 'Style inheritance depth limit exceeded.');
	visited.add(key);
	const parent = style.attributes.get(category);
	const inherited =
		parent !== undefined && parent !== id
			? styleSheet(parent, category, context, visited)
			: emptySheet();
	const full = mergeSheets(inherited, style.sheet);
	const pattern =
		category === 'LineStyle'
			? /^(Line|BeginArrow|EndArrow|Rounding|QuickStyleLine)/
			: category === 'FillStyle'
				? /^(Fill|Shdw|QuickStyleFill|QuickStyleEffectsMatrix|RotateGradientWithShape|UseGroupGradient)/
				: /^(LeftMargin|RightMargin|TopMargin|BottomMargin|VerticalAlign|TextBkgnd|TextBkgndTrans|QuickStyleFont.*)$/;
	const themeSelectors =
		/^(ThemeIndex|ColorSchemeIndex|EffectSchemeIndex|ConnectorSchemeIndex|FontSchemeIndex|VariationColorIndex|VariationStyleIndex|QuickStyleType|QuickStyleVariation)$/;
	const result = {
		cells: new Map(
			[...full.cells].filter(([key]) => pattern.test(key) || themeSelectors.test(key)),
		),
		sections: new Map(
			[...full.sections].filter(
				([, section]) =>
					(category === 'TextStyle' && ['Character', 'Paragraph', 'Tabs'].includes(section.name)) ||
					(category === 'FillStyle' && section.name === 'FillGradient') ||
					(category === 'LineStyle' && section.name === 'LineGradient'),
			),
		),
	};
	context.styleCache.set(key, result);
	return result;
}
