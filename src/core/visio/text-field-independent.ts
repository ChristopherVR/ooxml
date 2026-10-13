import { attribute } from './sheet';
import { VISIO_FIELD_FUNCTIONS } from './text-fields';

/** These exact native Field formulas contain no shape/text references. This is a
 * dependency proof, not permission to format their internal picture or evaluate dates. */
export function independentTextFieldFormula(cell: Element, formula: string): boolean {
	const row = cell.parentNode as Element | null,
		section = row?.parentNode as Element | null;
	if (
		cell.localName !== 'Cell' ||
		row?.localName !== 'Row' ||
		section?.localName !== 'Section' ||
		attribute(section, 'N') !== 'Field' ||
		row.hasAttribute('Del') ||
		section.hasAttribute('Del') ||
		row.hasAttribute('N') ||
		row.hasAttribute('T') ||
		!/^(0|[1-9]\d*)$/.test(attribute(row, 'IX') ?? '0') ||
		(attribute(section, 'IX') ?? '0') !== '0' ||
		cell.hasAttribute('E')
	)
		return false;
	if (attribute(cell, 'N') === 'Format')
		return /^\s*FIELDPICTURE\(\s*(2|37|204)\s*\)\s*$/i.test(formula);
	if (attribute(cell, 'N') !== 'Value') return false;
	if (/^\s*DATE\(\s*\d{1,4}\s*,\s*\d{1,2}\s*,\s*\d{1,2}\s*\)\s*$/i.test(formula)) return true;
	const direct = /^\s*([A-Z]+)\(\s*\)\s*$/i.exec(formula)?.[1]?.toUpperCase();
	return direct !== undefined && VISIO_FIELD_FUNCTIONS.includes(direct);
}
