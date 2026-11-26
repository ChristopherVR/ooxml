import { elements, type XmlElement } from '../../xml/index';
import type { PivotCalculatedField, PivotTable } from '../model';
import { decodeEscapes } from '../read/xml-util';

const identifier = /^[A-Za-z_][A-Za-z0-9_]{0,254}$/;

export function supportsPivotCalculatedField(pivot: PivotTable): boolean {
	return !(
		pivot.columnField ||
		pivot.pageFilter ||
		pivot.valuesAxis === 'rows' ||
		pivot.condition ||
		pivot.fieldGroups?.length ||
		pivot.dateGroups?.length ||
		pivot.fieldSubtotals?.some(({ functions }) => functions.length) ||
		(pivot.additionalValues?.length ?? 0) >= 15
	);
}

export function validatePivotCalculatedField(pivot: PivotTable): void {
	const calculated = pivot.calculatedField;
	if (!calculated) return;
	const caption = calculated.caption.toLowerCase();
	if (
		!supportsPivotCalculatedField(pivot) ||
		!identifier.test(calculated.name) ||
		!identifier.test(calculated.field) ||
		!pivot.fields.includes(calculated.field) ||
		pivot.fields.some((name) => name.toLowerCase() === calculated.name.toLowerCase()) ||
		!calculated.caption ||
		calculated.caption.length > 255 ||
		[
			pivot.valueCaption,
			...pivot.fields,
			calculated.name,
			...(pivot.additionalValues ?? []).map((v) => v.caption),
		].some((name) => name.toLowerCase() === caption) ||
		!Number.isFinite(calculated.factor) ||
		pivot.dateFields?.includes(calculated.field)
	)
		throw new Error(
			'Choose a unique calculated field name, a source field and a finite multiplier.',
		);
}

export function readPivotCalculatedField(
	cacheField: XmlElement,
	dataField: XmlElement,
): PivotCalculatedField {
	const formula = cacheField.getAttribute('formula') ?? '';
	const parts = /^([A-Za-z_][A-Za-z0-9_]{0,254})\*(-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?)$/.exec(
		formula,
	);
	const name = decodeEscapes(cacheField.getAttribute('name') ?? '');
	const caption = decodeEscapes(dataField.getAttribute('name') ?? '');
	if (
		!parts ||
		!name ||
		!caption ||
		cacheField.getAttribute('databaseField') !== '0' ||
		cacheField.getAttribute('numFmtId') !== '0' ||
		elements(cacheField).length ||
		(dataField.getAttribute('subtotal') ?? 'sum') !== 'sum' ||
		(dataField.getAttribute('showDataAs') ?? 'normal') !== 'normal'
	)
		throw new Error('Unsupported calculated PivotTable field.');
	return { name, field: parts[1]!, factor: Number(parts[2]), caption };
}
