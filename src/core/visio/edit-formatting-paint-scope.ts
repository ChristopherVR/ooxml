import type { VisioPackage } from './package';
import {
	attribute,
	child,
	children,
	readSheet,
	emptySheet,
	mergeSheets,
	number,
	type Report,
	type Cells,
	type Sheet,
} from './sheet';
import {
	effectiveShapeCell,
	assertEditableFormattingCell,
	uniqueFormattingCells,
} from './edit-style-admission';
import { executableCellFormula } from './cell-formula';
import { analyzeVisioFormula, evaluateVisioFormula, visioFormulaCachedValue } from './formula';
import { related, visioXml } from './parts';
import { styleSheet, type StyleRecord } from './style-inheritance';
import { loadVisioThemes } from './theme';
import { themeFormat } from './theme-resolve';
import { rootThemeSheet } from './theme-root';
import { themeChild } from './theme-color';
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
function scalar(cell: Element | undefined): number | undefined {
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

/** Do not silently alter saved gradient stops through foreground transparency edits.
 * Native theme paint selection reuses the parser's style and resource resolvers.
 */
export async function assertNonGradientPaint(
	pkg: VisioPackage,
	document: Element,
	shape: Element,
	pageId: string,
	paint: 'Fill' | 'Line',
	check: () => void,
): Promise<void> {
	const category = paint === 'Fill' ? 'FillStyle' : 'LineStyle';
	const pattern = effectiveShapeCell(shape, document, `${paint}Pattern`, category);
	if (pattern && attribute(pattern, 'V') !== 'Themed') {
		const value = scalar(pattern);
		if (value === 0) return;
		if (paint === 'Fill' && value !== undefined && value >= 25)
			fail('UNSUPPORTED_FORMAT_EDIT', 'Classic gradient and custom fill edits are unsupported.');
	}
	const enabled = effectiveShapeCell(shape, document, `${paint}GradientEnabled`, category);
	if (!enabled) return;
	if (attribute(enabled, 'V') !== 'Themed') {
		if (scalar(enabled) === 0) return;
		fail(
			'UNSUPPORTED_FORMAT_EDIT',
			'Gradient paint transparency and background edits are unsupported.',
		);
	}
	assertEditableFormattingCell(enabled);
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
	const resolved = number(sheet.cells, `${paint}GradientEnabled`, NaN);
	if (resolved === 0) {
		const source = styles.get('0');
		if (!source || root?.attributes.get('NameU') !== 'No Style')
			fail('UNSUPPORTED_FORMAT_EDIT', 'Root gradient context cannot be proven.');
		const node = children(child(document, 'StyleSheets'), 'StyleSheet').find(
			(item) => attribute(item, 'ID') === '0',
		);
		if (scalar(node ? uniqueFormattingCells(node).get(`${paint}GradientEnabled`) : undefined) !== 0)
			fail('UNSUPPORTED_FORMAT_EDIT', 'Root gradient state cannot be proven inactive.');
		return;
	}
	const selected = themeFormat(sheet.cells, paint, resources);
	if (paint === 'Fill' && selected && ['solidFill', 'noFill'].includes(selected.localName)) return;
	if (
		paint === 'Line' &&
		selected &&
		!themeChild(selected, 'gradFill') &&
		(themeChild(selected, 'solidFill') || themeChild(selected, 'noFill'))
	)
		return;
	fail(
		'UNSUPPORTED_FORMAT_EDIT',
		'Active or unresolved theme gradients cannot be edited through paint transparency.',
	);
}
