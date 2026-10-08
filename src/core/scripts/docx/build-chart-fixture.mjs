// Builds src/core/docx/__fixtures__/chart.docx: a Word package carrying one inline chart.
//
// Provenance: the chart part, its chart style and colour style parts and the theme are copied
// byte for byte from src/core/xlsx/__fixtures__/excel-features.xlsx (an Excel-authored clustered
// column chart, "Sales by region", two series of four regions). The chart part's relationships
// keep their ids and targets. Everything else (document.xml, rels, content types) is hand-built to
// look like what Word writes; there is no embedded workbook, so the chart is drawn from the values
// cached in the part. Word itself did not produce this package; open it in Word before treating it
// as evidence of Word compatibility.
//
// Run: node scripts/docx/build-chart-fixture.mjs
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const source = path.join(root, 'xlsx/__fixtures__/excel-features.xlsx');
const output = path.join(root, 'docx/__fixtures__/chart.docx');

const xlsx = await JSZip.loadAsync(await readFile(source));
const text = async (name) => {
	const file = xlsx.file(name);
	if (!file) throw new Error(`${name} is not in the source workbook`);
	return file.async('string');
};

const BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATE = new Date('2026-10-08T00:00:00Z'); // fixed, so the archive is reproducible
const zip = new JSZip();
const put = (name, content) => zip.file(name, content, { date: DATE });

put('word/theme/theme1.xml', await text('xl/theme/theme1.xml'));
put('word/charts/chart1.xml', await text('xl/charts/chart1.xml'));
put('word/charts/style1.xml', await text('xl/charts/style1.xml'));
put('word/charts/colors1.xml', await text('xl/charts/colors1.xml'));
put('word/charts/_rels/chart1.xml.rels', await text('xl/charts/_rels/chart1.xml.rels'));

const NS =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const chart = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="3200400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Chart 1" descr="Sales and cost by region"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"/></a:graphicData></a:graphic></wp:inline></w:drawing>`;
const paragraph = (inner) => `<w:p><w:r>${inner}</w:r></w:p>`;
const body = [
	paragraph('<w:t>Chart fixture. The chart below is inline:</w:t>'),
	paragraph(`<w:rPr><w:noProof/></w:rPr>${chart}`),
	paragraph('<w:t>End of document.</w:t>'),
].join('');
put(
	'word/document.xml',
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
);
put(
	'word/_rels/document.xml.rels',
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${BASE}/theme" Target="theme/theme1.xml"/><Relationship Id="rId2" Type="${BASE}/chart" Target="charts/chart1.xml"/></Relationships>`,
);
const CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
put(
	'[Content_Types].xml',
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="${CHART}"/><Override PartName="/word/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/><Override PartName="/word/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/></Types>`,
);
put(
	'_rels/.rels',
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${BASE}/officeDocument" Target="word/document.xml"/></Relationships>`,
);

await mkdir(path.dirname(output), { recursive: true });
await writeFile(
	output,
	await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', platform: 'UNIX' }),
);
console.log(`wrote ${path.relative(root, output)}`);
