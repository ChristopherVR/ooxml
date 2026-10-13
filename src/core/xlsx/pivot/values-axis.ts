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
