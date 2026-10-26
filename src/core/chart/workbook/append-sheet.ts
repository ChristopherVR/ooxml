// Allocate fresh chart data without rewriting existing worksheet cells or unknown package parts.
import { NS, buildXml, children, first, parseXml, type XmlElement } from '../../xml';
import { chartWorkbookRowsXml } from './create';
import { openEmbeddedPackage } from './writer';

export interface AppendChartWorkbookSheetOptions {
	nameBase?: string;
}
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
function element(parent: XmlElement, local: string): XmlElement {
	return parent.ownerDocument!.createElementNS(
		parent.namespaceURI,
		parent.prefix ? `${parent.prefix}:${local}` : local,
	);
}

/** Append one fresh, visible worksheet and its relationship/content type, retaining every other part. */
export async function appendChartWorkbookSheet(
	bytes: Uint8Array,
	rowsForName: (sheetName: string) => readonly (readonly (string | number | null)[])[],
	options: AppendChartWorkbookSheetOptions = {},
): Promise<{ bytes: Uint8Array; sheetName: string }> {
	const zip = await openEmbeddedPackage(bytes);
	const read = async (path: string) => {
		const text = await zip.file(path)?.async('string');
		if (!text) throw new Error(`Chart workbook part is missing: ${path}`);
		return parseXml(text, { label: 'Chart workbook' }).documentElement;
	};
	const workbook = await read('xl/workbook.xml'),
		rels = await read('xl/_rels/workbook.xml.rels'),
		types = await read('[Content_Types].xml');
	const protection = first(workbook, 'workbookProtection', NS.x);
	if (['1', 'true'].includes(protection?.getAttribute('lockStructure') ?? ''))
		throw new Error('The chart workbook structure is protected.');
	const sheets = first(workbook, 'sheets', NS.x);
	if (!sheets) throw new Error('The chart workbook has no sheets element.');
	const entries = children(sheets, 'sheet', NS.x);
	const names = new Set(entries.map((sheet) => (sheet.getAttribute('name') ?? '').toLowerCase()));
	const base = options.nameBase ?? 'ChartData';
	if (!base.trim() || /[\\/?*\[\]:]/u.test(base) || base.startsWith("'") || base.endsWith("'"))
		throw new RangeError('Invalid chart data sheet name.');
	let sheetName = '';
	for (let n = 1; ; n++) {
		const suffix = String(n);
		const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
		if (!names.has(candidate.toLowerCase())) {
			sheetName = candidate;
			break;
		}
	}
	const rows = rowsForName(sheetName);
	if (rows.length > 1048576 || rows.some((row) => row.length > 16384))
		throw new RangeError('Chart data exceeds worksheet bounds.');
	for (const row of rows)
		for (const value of row)
			if (typeof value === 'number' && !Number.isFinite(value))
				throw new RangeError('Chart workbook numbers must be finite.');
	const taken = new Set(Object.keys(zip.files).map((name) => name.toLowerCase()));
	let part = '';
	for (let n = 1; ; n++) {
		const candidate = `xl/worksheets/sheet${n}.xml`;
		if (!taken.has(candidate.toLowerCase())) {
			part = candidate;
			break;
		}
	}
	const ids = new Set(children(rels, 'Relationship', NS.rels).map((rel) => rel.getAttribute('Id')));
	let id = '';
	for (let n = 1; ; n++) {
		if (!ids.has(`rId${n}`)) {
			id = `rId${n}`;
			break;
		}
	}
	const sheetIds = entries.map((sheet) => Number(sheet.getAttribute('sheetId')));
	if (sheetIds.some((n) => !Number.isSafeInteger(n) || n < 1))
		throw new Error('The chart workbook has invalid sheet identifiers.');
	const sheet = element(sheets, 'sheet');
	sheet.setAttribute('name', sheetName);
	sheet.setAttribute('sheetId', String(Math.max(0, ...sheetIds) + 1));
	sheet.setAttributeNS(NS.r, 'r:id', id);
	sheets.appendChild(sheet);
	const relationship = element(rels, 'Relationship');
	relationship.setAttribute('Id', id);
	relationship.setAttribute('Type', `${NS.r}/worksheet`);
	relationship.setAttribute('Target', part.slice(3));
	rels.appendChild(relationship);
	const override = element(types, 'Override');
	override.setAttribute('PartName', `/${part}`);
	override.setAttribute(
		'ContentType',
		'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
	);
	types.appendChild(override);
	zip.file('xl/workbook.xml', DECLARATION + buildXml(workbook), { createFolders: false });
	zip.file('xl/_rels/workbook.xml.rels', DECLARATION + buildXml(rels), { createFolders: false });
	zip.file('[Content_Types].xml', DECLARATION + buildXml(types), { createFolders: false });
	zip.file(part, chartWorkbookRowsXml(rows), { createFolders: false });
	return {
		bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
		sheetName,
	};
}
