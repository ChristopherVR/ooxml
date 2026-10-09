import type { VisioPackage } from './package';
import {
	attribute,
	child,
	children,
	readSheet,
	emptySheet,
	mergeSheets,
	type Report,
	type Cells,
	type Sheet,
} from './sheet';
import { effectiveShapeCell, uniqueFormattingCells } from './edit-style-admission';
import { executableCellFormula } from './cell-formula';
import { analyzeVisioFormula, evaluateVisioFormula, visioFormulaCachedValue } from './formula';
import { related, visioXml } from './parts';
import { styleSheet, type StyleRecord } from './style-inheritance';
import { loadVisioThemes } from './theme';
import type { ThemeResources } from './theme-resolve';
import { rootThemeSheet } from './theme-root';
import { fail } from './package-common';

const selectors = [
	'ColorSchemeIndex',
	'EffectSchemeIndex',
	'ConnectorSchemeIndex',
	'VariationStyleIndex',
	'VariationColorIndex',
	'QuickStyleType',
	'QuickStyleVariation',
	'QuickStyleFillMatrix',
	'QuickStyleLineMatrix',
	'OneD',
];
export function scalar(cell: Element | undefined): number | undefined {
	if (!cell) return undefined;
	if (cell.hasAttribute('E')) fail('UNSUPPORTED_FORMAT_EDIT', 'Paint state has an error cache.');
	const cached = visioFormulaCachedValue(attribute(cell, 'V') ?? '', attribute(cell, 'U'));
	if (cached.unit !== 'scalar') fail('EDIT_FORMULA_UNIT', 'Paint state requires scalar units.');
	const source = executableCellFormula(attribute(cell, 'F'));
	if (source) {
		const analysis = analyzeVisioFormula(source);
		if (analysis.dynamic || analysis.references.length || analysis.unsupportedFunctions.length)
			fail('UNSUPPORTED_FORMAT_EDIT', 'Paint state dependencies cannot be proven.');
		const actual = evaluateVisioFormula(source, () =>
			fail('UNSUPPORTED_FORMAT_EDIT', 'Unknown paint dependency.'),
		);
		if (actual.unit !== 'scalar' || actual.value !== cached.value)
			fail('UNSUPPORTED_FORMAT_EDIT', 'Paint state formula cache is stale.');
	}
	return cached.value;
}

export interface ShapeThemeContext {
	/** The shape's merged style and local cells, with root theme selection applied. */
	sheet: Sheet;
	resources: ThemeResources;
	styles: Map<string, StyleRecord>;
}
/** The parser's theme resources for one local shape, refusing anything it cannot prove. */
export async function shapeThemeContext(
	pkg: VisioPackage,
	document: Element,
	shape: Element,
	pageId: string,
	check: () => void,
): Promise<ShapeThemeContext> {
	const report: Report = () =>
		fail('UNSUPPORTED_FORMAT_EDIT', 'Theme paint context cannot be resolved safely.');
	if (children(document, 'StyleSheets').length > 1)
		fail('UNSUPPORTED_FORMAT_EDIT', 'Ambiguous style containers cannot supply theme paint.');
	const styles = new Map<string, StyleRecord>();
	for (const container of children(document, 'StyleSheets'))
		for (const node of children(container, 'StyleSheet')) {
			check();
			const id = attribute(node, 'ID');
			if (!id || styles.has(id)) fail('UNSUPPORTED_FORMAT_EDIT', 'Ambiguous style IDs.');
			uniqueFormattingCells(node);
			styles.set(id, {
				sheet: readSheet(node),
				attributes: new Map(Array.from(node.attributes).map((item) => [item.name, item.value])),
			});
		}
	const context = { styles, styleCache: new Map<string, Sheet>(), report, checkTime: check };
	let sheet = emptySheet();
	for (const name of ['LineStyle', 'FillStyle', 'TextStyle'] as const) {
		for (const selector of selectors) scalar(effectiveShapeCell(shape, document, selector, name));
		const id =
			attribute(shape, name) ??
			attribute(child(document, 'DocumentSheet'), name) ??
			(styles.has('0') ? '0' : undefined);
		if (id !== undefined) sheet = mergeSheets(sheet, styleSheet(id, name, context));
	}
	sheet = mergeSheets(sheet, readSheet(shape));
	const documentPath = (await related(pkg, '', 'document'))!;
	const pagesPath = (await related(pkg, documentPath, 'pages'))!;
	const page = children(await visioXml(pkg, pagesPath, 'Pages'), 'Page').find(
		(node) => attribute(node, 'ID') === pageId,
	);
	const pageSheet = child(page, 'PageSheet');
	const pageCells: Cells = new Map();
	if (pageSheet) uniqueFormattingCells(pageSheet);
	for (const name of ['LineStyle', 'FillStyle', 'TextStyle'] as const) {
		const id = attribute(pageSheet, name);
		if (id !== undefined && pageSheet) {
			for (const selector of selectors)
				scalar(effectiveShapeCell(pageSheet, document, selector, name));
			for (const [key, value] of styleSheet(id, name, context).cells)
				if (
					/^(ColorSchemeIndex|EffectSchemeIndex|ConnectorSchemeIndex|FontSchemeIndex|VariationColorIndex|VariationStyleIndex)$/.test(
						key,
					)
				)
					pageCells.set(key, value);
		}
	}
	for (const node of children(pageSheet, 'Cell'))
		if (selectors.includes(attribute(node, 'N') ?? '')) scalar(node);
	const root = styles.get('0');
	const resources = {
		themes: await loadVisioThemes(pkg, documentPath, report),
		pageCells: new Map([...pageCells, ...readSheet(pageSheet).cells]),
		...(root?.attributes.get('NameU') === 'No Style' ? { rootSheet: root.sheet } : {}),
	};
	sheet = rootThemeSheet(sheet, resources, report);
	return { sheet, resources, styles };
}
