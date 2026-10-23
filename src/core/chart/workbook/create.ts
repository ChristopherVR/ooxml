/**
 * A new embedded workbook for a chart: the package Word and PowerPoint embed behind a chart they
 * create (`c:externalData`), holding its data sheet in the standard layout (`layout.ts`), so
 * "Edit Data" in Office opens the chart's own numbers. One worksheet, inline strings, a minimal
 * style sheet; nothing else (no table part, no theme), which Excel opens without repair.
 *
 * @module chart/workbook/create
 */
import JSZip from 'jszip';
import type { ChartDataSheet } from '../data-sheet';
import { formatCellAddress } from './cellref';
import { DEFAULT_CHART_SHEET_NAME } from './layout';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const escape = (text: string) =>
	text
		.replace(/&/gu, '&amp;')
		.replace(/</gu, '&lt;')
		.replace(/>/gu, '&gt;')
		.replace(/"/gu, '&quot;');

function cell(col: number, row: number, value: string | number | null): string {
	const ref = formatCellAddress({ col, row });
	if (value === null || value === '') return '';
	if (typeof value === 'number')
		return Number.isFinite(value) ? `<c r="${ref}"><v>${value}</v></c>` : '';
	return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escape(value)}</t></is></c>`;
}

/** The worksheet part of `sheet` in the standard layout. */
export function chartDataWorksheetXml(sheet: ChartDataSheet): string {
	const rows: string[] = [];
	const header = sheet.series.map((series, index) => cell(index + 1, 0, series.name)).join('');
	rows.push(`<row r="1">${header}</row>`);
	sheet.categories.forEach((label, category) => {
		const values = sheet.series
			.map((series, index) => cell(index + 1, category + 1, series.values[category] ?? null))
			.join('');
		rows.push(`<row r="${category + 2}">${cell(0, category + 1, label)}${values}</row>`);
	});
	const last = formatCellAddress({
		col: Math.max(1, sheet.series.length),
		row: Math.max(1, sheet.categories.length),
	});
	return `${DECLARATION}<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><dimension ref="A1:${last}"/><sheetData>${rows.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
}

const STYLES = `${DECLARATION}<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

/** A new `.xlsx` package holding `sheet` on one worksheet named `sheetName`. */
export async function createChartWorkbook(
	sheet: ChartDataSheet,
	sheetName = DEFAULT_CHART_SHEET_NAME,
): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`${DECLARATION}<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
	);
	zip.file(
		'xl/workbook.xml',
		`${DECLARATION}<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="${escape(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
	);
	zip.file(
		'xl/_rels/workbook.xml.rels',
		`${DECLARATION}<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
	);
	zip.file('xl/worksheets/sheet1.xml', chartDataWorksheetXml(sheet));
	zip.file('xl/styles.xml', STYLES);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
