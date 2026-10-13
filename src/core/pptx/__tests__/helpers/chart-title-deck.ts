import JSZip from 'jszip';
import { PresentationBuilder } from 'ooxml-core/pptx';

/** Synthetic chart with independent rich title, a linked workbook and opaque title formatting. */
export async function createChartTitleDeck(linked = false): Promise<ArrayBuffer> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides = [
		createSlide('Blank')
			.addChart(
				'bar',
				{
					title: 'Annual report',
					categories: ['Q1', 'Q2'],
					series: [{ name: 'Dataset', values: [10, 20] }],
				},
				{ x: 80, y: 60, width: 700, height: 400 },
			)
			.build(),
	];
	if (linked)
		data.slides[0]!.elements.push(
			createSlide('Blank')
				.addText('Slide-only report', { x: 80, y: 480, width: 400, height: 40, fontSize: 20 })
				.build().elements[0]!,
		);
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();
	const title = linked
		? '<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Annual report</c:v></c:pt></c:strCache></c:strRef></c:tx><c:overlay val="0"/></c:title>'
		: '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"><a:defRPr sz="1800"/></a:pPr><a:r><a:rPr lang="fr-FR" sz="1800" b="1" kern="1200"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:latin typeface="Arial"/><a:hlinkClick r:id="rIdTitleLink"/></a:rPr><a:t xml:space="preserve">Annual re</a:t></a:r><a:r><a:rPr lang="en-GB" sz="1800" i="1" u="sng"/><a:t>port</a:t></a:r><a:endParaRPr lang="fr-FR"/></a:p></c:rich></c:tx><c:overlay val="0"/><c:extLst><c:ext uri="chart-title-parity"><parity:retained xmlns:parity="urn:parity:chart-title"/></c:ext></c:extLst></c:title>';
	const chartPath = 'ppt/charts/chart1.xml';
	let chart = await zip.file(chartPath)!.async('string');
	chart = chart.replace(/<c:title>[\s\S]*?<\/c:title>/u, title);
	chart = chart.replace(
		'</c:chartSpace>',
		'<c:externalData r:id="rIdWorkbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>',
	);
	zip.file(chartPath, chart);
	zip.file(
		'ppt/charts/_rels/chart1.xml.rels',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Chart_Data.xlsx"/><Relationship Id="rIdTitleLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/chart-title" TargetMode="External"/></Relationships>',
	);
	const workbook = new JSZip();
	workbook.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
	);
	workbook.file(
		'_rels/.rels',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
	);
	workbook.file(
		'xl/workbook.xml',
		'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
	);
	workbook.file(
		'xl/_rels/workbook.xml.rels',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
	);
	workbook.file(
		'xl/worksheets/sheet1.xml',
		'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Annual report</t></is></c><c r="B1" t="inlineStr"><is><t>Dataset</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c><c r="B3"><v>20</v></c></row></sheetData></worksheet>',
	);
	zip.file('ppt/embeddings/Chart_Data.xlsx', await workbook.generateAsync({ type: 'uint8array' }));
	zip.file(
		'[Content_Types].xml',
		(await zip.file('[Content_Types].xml')!.async('string')).replace(
			'</Types>',
			'<Override PartName="/ppt/embeddings/Chart_Data.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>',
		),
	);
	return zip.generateAsync({ type: 'arraybuffer' });
}
