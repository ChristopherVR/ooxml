import { isCellError, type Cell, type Workbook } from './model';
import { formatValue } from './numfmt';
import { styleAt } from './styles';

/** The formatted text used by value filters and editing commands. */
export function displayText(workbook: Workbook, cell: Cell | undefined): string {
	if (!cell || cell.value === null) return '';
	const value = cell.value;
	if (isCellError(value)) return value.error;
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (cell.richText && typeof value === 'string') return value;
	return formatValue(value, styleAt(workbook, cell.styleId).numFmt, {
		date1904: workbook.date1904,
	}).text;
}
