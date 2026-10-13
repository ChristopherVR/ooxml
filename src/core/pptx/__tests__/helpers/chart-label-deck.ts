import JSZip from 'jszip';

import { createChartTitleDeck } from './chart-title-deck';

/** Synthetic chart whose caches address the matching literal workbook cells. */
export async function createChartLabelDeck(
	options: {
		formula?: boolean;
		external?: boolean;
		missing?: boolean;
		inconsistent?: boolean;
		shared?: boolean;
	} = {},
): Promise<ArrayBuffer> {
	const zip = await JSZip.loadAsync(await createChartTitleDeck());
	const path = 'ppt/charts/chart1.xml';
	let chart = await zip.file(path)!.async('string');
	chart = chart.replace(
		/<c:tx><c:v>Dataset<\/c:v><\/c:tx>/u,
		'<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Dataset</c:v></c:pt></c:strCache></c:strRef></c:tx>',
	);
	chart = chart.replace(
		/<c:cat><c:strLit>([\s\S]*?)<\/c:strLit><\/c:cat>/u,
		'<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache>$1</c:strCache></c:strRef></c:cat>',
	);
	chart = chart.replace(
		/<c:val><c:numLit>([\s\S]*?)<\/c:numLit><\/c:val>/u,
		'<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache>$1</c:numCache></c:numRef></c:val>',
	);
	zip.file(path, chart);
	if (options.external)
		zip.file(
			'ppt/charts/_rels/chart1.xml.rels',
			(await zip.file('ppt/charts/_rels/chart1.xml.rels')!.async('string')).replace(
				'Id="rIdWorkbook"',
				'Id="rIdWorkbook" TargetMode="External"',
			),
		);
	if (options.shared)
		zip.file(
			'ppt/charts/_rels/other.xml.rels',
			'<Relationships><Relationship Id="r1" Target="../embeddings/Chart_Data.xlsx"/></Relationships>',
		);
	const embedding = 'ppt/embeddings/Chart_Data.xlsx';
	if (options.missing) zip.remove(embedding);
	else {
		const workbook = await JSZip.loadAsync(await zip.file(embedding)!.async('uint8array'));
		let sheet = await workbook.file('xl/worksheets/sheet1.xml')!.async('string');
		if (options.formula) sheet = sheet.replace('<v>10</v>', '<f>5+5</f><v>10</v>');
		if (options.inconsistent) sheet = sheet.replace('<t>Dataset</t>', '<t>Different</t>');
		sheet = sheet.replace(
			'</sheetData>',
			'<row r="20"><c r="Z20" s="7" t="inlineStr"><is><t xml:space="preserve">001 untouched </t></is></c></row></sheetData>',
		);
		workbook.file('xl/worksheets/sheet1.xml', sheet);
		workbook.file('unrelated.xml', '<retained value="opaque"/>');
		workbook.file(
			'xl/styles.xml',
			'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="8">' +
				'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'.repeat(8) +
				'</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
		);
		workbook.file(
			'[Content_Types].xml',
			(await workbook.file('[Content_Types].xml')!.async('string')).replace(
				'</Types>',
				'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
			),
		);
		workbook.file(
			'xl/_rels/workbook.xml.rels',
			(await workbook.file('xl/_rels/workbook.xml.rels')!.async('string')).replace(
				'</Relationships>',
				'<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
			),
		);

		zip.file(embedding, await workbook.generateAsync({ type: 'uint8array' }));
	}
	return zip.generateAsync({ type: 'arraybuffer' });
}
