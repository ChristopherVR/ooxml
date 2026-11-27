// Allocate fresh chart data without rewriting existing worksheet cells or unknown package parts.
import { resolvePartPath } from '../../opc/relationships';
import { dateZipEntries } from '../../opc/reproducible';
import { NS, buildXml, children, first, parseXml, type XmlElement } from '../../xml';
import { chartWorkbookRowsXml } from './create';
import { openEmbeddedPackage } from './writer';

export interface AppendChartWorkbookSheetOptions {
	nameBase?: string;
	/**
	 * A sheet to write the rows over, keeping its name and place, instead of adding one: chart data
	 * the caller allocated earlier that nothing else reads. A name the workbook lacks adds a sheet.
	 */
	reuse?: string;
	/**
	 * Sheets to remove with their parts: chart data the caller allocated that nothing reads any more.
	 * The sheet the rows are written to is kept, and names the workbook lacks are ignored.
	 */
	drop?: readonly string[];
}
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const WORKBOOK = 'xl/workbook.xml';
function element(parent: XmlElement, local: string): XmlElement {
	return parent.ownerDocument!.createElementNS(
		parent.namespaceURI,
		parent.prefix ? `${parent.prefix}:${local}` : local,
	);
}
const freeName = (taken: (candidate: string) => boolean, make: (n: number) => string) => {
	for (let n = 1; ; n++) if (!taken(make(n))) return make(n);
};
const relsPartFor = (part: string) => part.replace(/([^/]+)$/u, '_rels/$1.rels');

/**
 * Write chart rows to one visible worksheet, a fresh one or `options.reuse`, and select it; remove
 * `options.drop`. Every other sheet, cell and package part is kept.
 */
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
	const workbook = await read(WORKBOOK),
		rels = await read('xl/_rels/workbook.xml.rels'),
		types = await read('[Content_Types].xml');
	const protection = first(workbook, 'workbookProtection', NS.x);
	if (['1', 'true'].includes(protection?.getAttribute('lockStructure') ?? ''))
		throw new Error('The chart workbook structure is protected.');
	const sheets = first(workbook, 'sheets', NS.x);
	if (!sheets) throw new Error('The chart workbook has no sheets element.');
	const entries = children(sheets, 'sheet', NS.x);
	const sheetIds = entries.map((sheet) => Number(sheet.getAttribute('sheetId')));
	if (sheetIds.some((n) => !Number.isSafeInteger(n) || n < 1))
		throw new Error('The chart workbook has invalid sheet identifiers.');
	const relationships = children(rels, 'Relationship', NS.rels);
	const overrides = children(types, 'Override', NS.contentTypes);
	const named = (name: string | undefined) =>
		name === undefined
			? undefined
			: entries.find((sheet) => sheet.getAttribute('name')?.toLowerCase() === name.toLowerCase());
	/** The worksheet part and relationship of `sheet`, when it has an internal one. */
	const partOf = (sheet: XmlElement) => {
		const rel = relationships.find(
			(item) => item.getAttribute('Id') === sheet.getAttributeNS(NS.r, 'id'),
		);
		const target = rel?.getAttribute('Target');
		return rel && target && rel.getAttribute('TargetMode') !== 'External'
			? { rel, part: resolvePartPath(WORKBOOK, target) }
			: undefined;
	};
	const reused = named(options.reuse);
	const reusedPart = reused && partOf(reused);
	let target: XmlElement, part: string;
	if (reused && reusedPart) {
		target = reused;
		part = reusedPart.part;
	} else {
		const names = new Set(entries.map((sheet) => (sheet.getAttribute('name') ?? '').toLowerCase()));
		const base = options.nameBase ?? 'ChartData';
		if (!base.trim() || /[\\/?*\[\]:]/u.test(base) || base.startsWith("'") || base.endsWith("'"))
			throw new RangeError('Invalid chart data sheet name.');
		const sheetName = freeName(
			(candidate) => names.has(candidate.toLowerCase()),
			(n) => `${base.slice(0, 31 - String(n).length)}${n}`,
		);
		const taken = new Set(Object.keys(zip.files).map((name) => name.toLowerCase()));
		part = freeName(
			(candidate) => taken.has(candidate.toLowerCase()),
			(n) => `xl/worksheets/sheet${n}.xml`,
		);
		const ids = new Set(relationships.map((rel) => rel.getAttribute('Id')));
		const id = freeName(
			(candidate) => ids.has(candidate),
			(n) => `rId${n}`,
		);
		target = element(sheets, 'sheet');
		target.setAttribute('name', sheetName);
		target.setAttribute('sheetId', String(Math.max(0, ...sheetIds) + 1));
		target.setAttributeNS(NS.r, 'r:id', id);
		sheets.appendChild(target);
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
	}
	const sheetName = target.getAttribute('name')!;
	const rows = rowsForName(sheetName);
	if (rows.length > 1048576 || rows.some((row) => row.length > 16384))
		throw new RangeError('Chart data exceeds worksheet bounds.');
	rows.forEach((row) => {
		for (const value of row)
			if (typeof value === 'number' && !Number.isFinite(value))
				throw new RangeError('Chart workbook numbers must be finite.');
	});
	const removed: string[] = [];
	const definedNames = first(workbook, 'definedNames', NS.x);
	for (const name of options.drop ?? []) {
		const sheet = named(name);
		const dropped = sheet && sheet !== target && partOf(sheet);
		if (!sheet || !dropped || !sheet.parentNode) continue;
		// Sheet-scoped names follow their sheet; the ones after it move up one place.
		const index = children(sheets, 'sheet', NS.x).indexOf(sheet);
		for (const definedName of definedNames ? children(definedNames, 'definedName', NS.x) : []) {
			const scope = Number(definedName.getAttribute('localSheetId') ?? NaN);
			if (scope === index) definedNames!.removeChild(definedName);
			else if (scope > index) definedName.setAttribute('localSheetId', String(scope - 1));
		}
		sheets.removeChild(sheet);
		rels.removeChild(dropped.rel);
		for (const override of overrides)
			if (override.getAttribute('PartName')?.toLowerCase() === `/${dropped.part}`.toLowerCase())
				types.removeChild(override);
		removed.push(dropped.part, relsPartFor(dropped.part));
	}
	if (definedNames && children(definedNames, 'definedName', NS.x).length === 0)
		workbook.removeChild(definedNames);
	// Edit Data must open the sheet the chart now reads, while other worksheet parts stay intact.
	let views = first(workbook, 'bookViews', NS.x);
	if (!views) {
		views = element(workbook, 'bookViews');
		workbook.insertBefore(views, sheets);
	}
	let workbookViews = children(views, 'workbookView', NS.x);
	if (workbookViews.length === 0) {
		const view = element(views, 'workbookView');
		views.appendChild(view);
		workbookViews = [view];
	}
	const active = children(sheets, 'sheet', NS.x).indexOf(target);
	for (const view of workbookViews) {
		view.setAttribute('activeTab', String(active));
		if (Number(view.getAttribute('firstSheet') ?? 0) > active)
			view.setAttribute('firstSheet', String(active));
	}
	for (const path of removed) zip.remove(path);
	zip.file(WORKBOOK, DECLARATION + buildXml(workbook), { createFolders: false });
	zip.file('xl/_rels/workbook.xml.rels', DECLARATION + buildXml(rels), { createFolders: false });
	zip.file('[Content_Types].xml', DECLARATION + buildXml(types), { createFolders: false });
	zip.file(part, chartWorkbookRowsXml(rows), { createFolders: false });
	dateZipEntries(zip);
	return {
		bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
		sheetName,
	};
}
