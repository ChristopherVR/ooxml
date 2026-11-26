import type { CellView, PivotTable, ThemePalette } from '../../index';
import { pivotColumnFields, pivotRowFields } from '../../pivot/axes';
import { tableStyleColor } from '../commands/style-presets';

/** A preview of built-in styles. Explicit cell and conditional formats take precedence. */
export function paintPivotStyle(
	view: CellView,
	pivot: PivotTable,
	row: number,
	col: number,
	theme: ThemePalette,
): CellView {
	const style = pivot.display?.style;
	const match = /^PivotStyle(Light|Medium|Dark)(\d{1,2})$/.exec(style?.name ?? 'PivotStyleLight16');
	if (!match) return view;
	const family = match[1] as 'Light' | 'Medium' | 'Dark';
	const index = Number(match[2]);
	const y = row - pivot.location.start.row - (pivot.pageFilter ? 2 : 0);
	if (y < 0) return view;
	const x = col - pivot.location.start.col;
	const columns = pivotColumnFields(pivot).length;
	const onRows = pivot.valuesAxis === 'rows';
	const multi = !onRows && !!(pivot.additionalValues?.length || pivot.calculatedField);
	const headers = columns ? columns + 1 + Number(multi) : 1;
	const labels = pivot.layout === 'compact' ? 1 : pivotRowFields(pivot).length + Number(onRows);
	const header =
		(style?.showColHeaders !== false && y < headers) ||
		(style?.showRowHeaders !== false && x < labels && y >= headers);
	const band =
		(style?.showRowStripes && y >= headers && (y - headers) % 2 === 0) ||
		(style?.showColStripes && x >= labels && (x - labels) % 2 === 0);
	if (!header && !band) return view;
	const color = tableStyleColor({ name: match[0], family, index }, theme);
	const tint = header ? (family === 'Light' ? '29' : '48') : '20';
	if (header) view.font.bold = true;
	view.fill ??= { background: `#${color}${tint}` };
	return view;
}
