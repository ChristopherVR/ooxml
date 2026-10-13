import type { XmlElement } from '../../xml/index';
import { attrs } from '../write/xml-out';
export function pivotFieldIndex(field: XmlElement | undefined): number {
	const raw = field?.getAttribute('x');
	if (raw === null || raw === undefined || !/^-?\d+$/.test(raw.trim())) return NaN;
	const value = Number(raw);
	return Number.isSafeInteger(value) ? value : NaN;
}

/** Native between-field row Values carry an explicit position matching their axis slot. */
export function pivotValueIndex(root: XmlElement, axis: readonly XmlElement[]): number {
	const index = axis.findIndex((field) => pivotFieldIndex(field) === -2);
	const position = root.getAttribute('dataPosition');
	if (position !== null && (!/^\d+$/.test(position.trim()) || Number(position) !== index))
		throw new Error('Inconsistent Values axis position');
	if (axis.length === 3 && index === 1 && position === null)
		throw new Error('Missing between-field Values position');
	return index;
}
/** Native Values-on-columns entries, with each measure nested inside each column item. */
export function pivotValueAxisItems(measures: number, columns?: number): string {
	const entries: string[] = [];
	if (columns === undefined) {
		for (let measure = 0; measure < measures; measure++)
			entries.push(
				`<i${attrs({ i: measure || undefined })}><x${attrs({ v: measure || undefined })}/></i>`,
			);
	} else {
		for (let column = 0; column < columns; column++)
			for (let measure = 0; measure < measures; measure++) {
				entries.push(
					`<i${attrs({ r: measure ? 1 : undefined, i: measure || undefined })}>${measure ? '' : `<x${attrs({ v: column || undefined })}/>`}<x${attrs({ v: measure || undefined })}/></i>`,
				);
			}
		for (let measure = 0; measure < measures; measure++)
			entries.push(`<i${attrs({ t: 'grand', i: measure || undefined })}><x/></i>`);
	}
	return `<colItems count="${entries.length}">${entries.join('')}</colItems>`;
}

/** Values follow the column field, or sit last/between two real row fields. */
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
			(rows.length === 2 || rows.length === 3) &&
			index(rows[0]) >= 0 &&
			rows.filter((field) => index(field) === -2).length === 1 &&
			rows.filter((field) => index(field) !== -2).every((field) => index(field) >= 0) &&
			columns.length <= 1 &&
			(!columns.length || index(columns[0]) >= 0)
		);
	if (rows.some((field) => !Number.isInteger(index(field)) || index(field) < 0)) return false;
	return measures === 1
		? columns.length <= 1
		: columns.length >= 1 &&
				columns.length <= 2 &&
				index(columns.at(-1)) === -2 &&
				(columns.length !== 2 || index(columns[0]) >= 0);
}
