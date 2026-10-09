import type { VisioPackage } from './package';
import { attribute, child, children, number } from './sheet';
import {
	effectiveShapeCell,
	assertEditableFormattingCell,
	uniqueFormattingCells,
} from './edit-style-admission';
import { themeFormat } from './theme-resolve';
import { themeChild } from './theme-color';
import { fail } from './package-common';
import { scalar, shapeThemeContext } from './edit-formatting-theme-context';

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
	const { sheet, resources, styles } = await shapeThemeContext(pkg, document, shape, pageId, check);
	const root = styles.get('0');
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
