import { elements, type XmlElement } from '../../xml/index';
import type { PivotCalculatedField, PivotTable } from '../model';
import { att, decodeEscapes } from '../read/xml-util';

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

export function splitCalculatedCacheField(allFields: XmlElement[]) {
	const formulas = allFields.filter((field) => field.hasAttribute('formula'));
	const formula = formulas[0];
	if (formulas.length > 1 || (formula && formula !== allFields.at(-1)))
		throw new Error('Multiple calculated PivotTable fields are preserved.');
	return { groupedFields: formula ? allFields.slice(0, -1) : allFields, formula };
}

export function readPivotCalculatedField(
	formula: XmlElement | undefined,
	dataFields: XmlElement[],
	cacheFieldCount: number,
): PivotCalculatedField | undefined {
	if (!formula) return;
	const last = dataFields.at(-1);
	if (dataFields.length <= 1 || !last || att(last, 'fld') !== String(cacheFieldCount - 1))
		throw new Error('Unsupported calculated PivotTable field.');
	return parseCalculatedField(formula, last);
}

export function assertPivotFieldPlacement(
	tableFields: XmlElement[],
	calculatedIndex: number,
): void {
	if (
		tableFields.some((field, index) =>
			['dragToRow', 'dragToCol', 'dragToPage'].some(
				(option) =>
					field.hasAttribute(option) && (index !== calculatedIndex || att(field, option) !== '0'),
			),
		)
	)
		throw new Error('Unsupported PivotTable field placement restrictions.');
}

function parseCalculatedField(cacheField: XmlElement, dataField: XmlElement): PivotCalculatedField {
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
