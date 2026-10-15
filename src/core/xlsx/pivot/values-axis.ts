import type { XmlElement } from '../../xml/index';
import { attrs } from '../write/xml-out';
import type { PivotGrid } from './aggregate';

/**
 * Selected reports retain the full cache axis lists; hierarchy tuples use their item indexes.
 * The Values level (`null`) holds measure indexes and keeps them.
 */
export function remapPivotRowEntries(
	entries: NonNullable<PivotGrid['rowEntries']>,
	axes: readonly (readonly number[] | null)[],
): NonNullable<PivotGrid['rowEntries']> {
	return entries.map((entry) =>
		entry.type === 'grand'
			? entry
			: {
					...entry,
					indices: entry.indices.map((item, axis) => {
						if (axes[axis] === null) return item;
						const index = axes[axis]?.[item];
						if (index === undefined) throw new Error('Invalid report row axis item');
						return index;
					}),
				},
	);
}
export function pivotFieldIndex(field: XmlElement | undefined): number {
	const raw = field?.getAttribute('x');
	if (raw === null || raw === undefined || !/^-?\d+$/.test(raw.trim())) return NaN;
	const value = Number(raw);
	return Number.isSafeInteger(value) ? value : NaN;
}

/** Native row Values before the last field carry an explicit position matching their axis slot. */
export function pivotValueIndex(root: XmlElement, axis: readonly XmlElement[]): number {
	const index = axis.findIndex((field) => pivotFieldIndex(field) === -2);
	const position = root.getAttribute('dataPosition');
	if (position !== null && (!/^\d+$/.test(position.trim()) || Number(position) !== index))
		throw new Error('Inconsistent Values axis position');
	if (index >= 0 && index < axis.length - 1 && position === null)
		throw new Error('Missing leading or between-field Values position');
	return index;
}
/** Native Values-on-columns entries of a report without column fields: one per measure. */
export function pivotValueAxisItems(measures: number): string {
	const entries: string[] = [];
	for (let measure = 0; measure < measures; measure++)
		entries.push(
			`<i${attrs({ i: measure || undefined })}><x${attrs({ v: measure || undefined })}/></i>`,
		);
	return `<colItems count="${entries.length}">${entries.join('')}</colItems>`;
}

/** Values follow the column fields, or sit anywhere among the real row fields. */
export function supportedPivotAxes(
	rows: readonly XmlElement[],
	columns: readonly XmlElement[],
	measures: number,
	valuesOnRows: boolean,
): boolean {
	const index = pivotFieldIndex;
	if (valuesOnRows)
		return (
			measures > 1 &&
			rows.length > 1 &&
			rows.filter((field) => index(field) === -2).length === 1 &&
			rows.filter((field) => index(field) !== -2).every((field) => index(field) >= 0) &&
			columns.every((field) => index(field) >= 0)
		);
	if (rows.some((field) => !Number.isInteger(index(field)) || index(field) < 0)) return false;
	// Several measures put Values last on the columns, after any number of column fields.
	if (measures > 1 && (!columns.length || index(columns.at(-1)) !== -2)) return false;
	return (measures > 1 ? columns.slice(0, -1) : columns).every(
		(field) => Number.isInteger(index(field)) && index(field) >= 0,
	);
}
