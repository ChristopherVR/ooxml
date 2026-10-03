// Hand-written test packages for round-trip tests (the shape of the review's `mk.py`).
import JSZip from 'jszip';
import type { Workbook } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { saveXlsx } from './index.js';

export const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const DEFAULT_STYLES = `${HEAD}<styleSheet xmlns="${X}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

export interface MiniPackage {
	sheets: { name: string; xml: string; rels?: string }[];
	styles?: string;
	parts?: Record<string, string>;
	contentTypes?: string;
	workbookRels?: string;
}

/** A worksheet part: `before` and `after` surround `<sheetData>`. */
export const ws = (rows: string, before = '', after = '') =>
	`${HEAD}<worksheet xmlns="${X}" xmlns:r="${R}">${before}<sheetData>${rows}</sheetData>${after}</worksheet>`;

/** Builds a minimal hand-written package (the shape of the review's `mk.py`). */
export async function miniPackage(spec: MiniPackage): Promise<Uint8Array> {
	const zip = new JSZip();
	const overrides = spec.sheets
		.map(
			(_, i) =>
				`<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
		)
		.join('');
	zip.file(
		'[Content_Types].xml',
		`${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${overrides}${spec.contentTypes ?? ''}</Types>`,
	);
	zip.file(
		'_rels/.rels',
		`${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
	);
	const rels = spec.sheets
		.map(
			(_, i) =>
				`<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
		)
		.join('');
	zip.file(
		'xl/_rels/workbook.xml.rels',
		`${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}<Relationship Id="rIdS" Type="${R}/styles" Target="styles.xml"/>${spec.workbookRels ?? ''}</Relationships>`,
	);
	const sheets = spec.sheets
		.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
		.join('');
	zip.file(
		'xl/workbook.xml',
		`${HEAD}<workbook xmlns="${X}" xmlns:r="${R}"><sheets>${sheets}</sheets></workbook>`,
	);
	zip.file('xl/styles.xml', spec.styles ?? DEFAULT_STYLES);
	spec.sheets.forEach((s, i) => {
		zip.file(`xl/worksheets/sheet${i + 1}.xml`, s.xml);
		if (s.rels) zip.file(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, s.rels);
	});
	for (const [name, xml] of Object.entries(spec.parts ?? {})) zip.file(name, xml);
	return zip.generateAsync({ type: 'uint8array' });
}

/** Loads, optionally edits, saves, and returns the saved workbook and a part reader. */
export async function roundTrip(
	bytes: Uint8Array,
	edit?: (wb: Workbook) => void,
): Promise<{ wb: Workbook; again: Workbook; part: (name: string) => Promise<string | undefined> }> {
	const wb = await loadXlsx(bytes);
	edit?.(wb);
	const saved = await saveXlsx(wb);
	const zip = await JSZip.loadAsync(saved);
	return {
		wb,
		again: await loadXlsx(saved),
		part: async (name) => zip.file(name)?.async('string'),
	};
}
