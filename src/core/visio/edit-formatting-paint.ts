import type { VisioShapeFormatEdit } from './edit-formatting-commands';
import type { FormattingWrite } from './edit-formatting';
import { assertNonGradientPaint } from './edit-formatting-paint-scope';
import type { VisioPackage } from './package';

export function shapeFormattingWrites(edit: VisioShapeFormatEdit): FormattingWrite[] {
	const writes = new Map<string, FormattingWrite>();
	const add = (
		name: string,
		value: string | number,
		category: 'FillStyle' | 'LineStyle',
		unit?: string,
		formula?: string,
	) =>
		writes.set(name, {
			name,
			value: String(value),
			category,
			...(unit ? { unit } : {}),
			...(formula ? { formula } : {}),
		});
	if (edit.fillColor !== undefined) {
		add('FillPattern', edit.fillColor === 'none' ? 0 : 1, 'FillStyle');
		add('FillGradientEnabled', 0, 'FillStyle');
		if (edit.fillColor !== 'none') {
			add('FillForegnd', edit.fillColor, 'FillStyle', undefined, rgb(edit.fillColor));
			add('FillForegndTrans', 0, 'FillStyle');
		}
	}
	if (edit.fillPattern !== undefined) {
		add('FillPattern', edit.fillPattern, 'FillStyle');
		add('FillGradientEnabled', 0, 'FillStyle');
	}
	if (edit.fillBackgroundColor !== undefined) {
		add(
			'FillBkgnd',
			edit.fillBackgroundColor,
			'FillStyle',
			undefined,
			rgb(edit.fillBackgroundColor),
		);
	}
	if (edit.fillTransparency !== undefined) {
		for (const name of ['FillForegndTrans', 'FillBkgndTrans'])
			add(name, edit.fillTransparency / 100, 'FillStyle');
	}
	if (edit.lineColor !== undefined) {
		add('LineColor', edit.lineColor, 'LineStyle', undefined, rgb(edit.lineColor));
		add('LineColorTrans', 0, 'LineStyle');
		add('LineGradientEnabled', 0, 'LineStyle');
	}
	if (edit.linePattern !== undefined) add('LinePattern', edit.linePattern, 'LineStyle');
	if (edit.lineTransparency !== undefined) {
		add('LineColorTrans', edit.lineTransparency / 100, 'LineStyle');
	}
	if (edit.lineWeight !== undefined) add('LineWeight', edit.lineWeight / 72, 'LineStyle', 'PT');
	return [...writes.values()];
}
export async function assertShapeFormattingPaintScope(
	pkg: VisioPackage,
	document: Element,
	shape: Element,
	edit: VisioShapeFormatEdit,
	changed: ReadonlyMap<string, Element | undefined>,
	check: () => void,
): Promise<void> {
	if (
		edit.fillColor === undefined &&
		edit.fillPattern === undefined &&
		['FillBkgnd', 'FillForegndTrans', 'FillBkgndTrans'].some((name) => changed.has(name))
	)
		await assertNonGradientPaint(pkg, document, shape, edit.pageId, 'Fill', check);
	if (edit.lineColor === undefined && changed.has('LineColorTrans'))
		await assertNonGradientPaint(pkg, document, shape, edit.pageId, 'Line', check);
}
function rgb(color: string): string {
	return `RGB(${[1, 3, 5].map((index) => parseInt(color.slice(index, index + 2), 16)).join(',')})`;
}
