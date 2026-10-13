import type { XmlElement } from '../../xml/index';
import { attrs } from '../write/xml-out';
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

/** Values are modeled last on their axis; row placement has one real tabular row field. */
export function supportedPivotAxes(
	rows: readonly XmlElement[],
	columns: readonly XmlElement[],
	measures: number,
	valuesOnRows: boolean,
): boolean {
	const index = (field: XmlElement | undefined) => {
		const raw = field?.getAttribute('x');
		const value = raw === null || raw === undefined || !raw.trim() ? NaN : Number(raw);
		return Number.isInteger(value) ? value : NaN;
	};
	if (valuesOnRows)
		return (
			measures > 1 &&
			rows.length === 2 &&
			index(rows[0]) >= 0 &&
			index(rows[1]) === -2 &&
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
