// The cell-floats documents: small hand-written packages with floating objects (`wp:anchor`)
// anchored in a table cell, and floating charts, which `scripts/docx-word-cell-float-probe.ps1`
// renders with Word and `word-positions.json` records. Tests build the same packages from here.
//
// Courier New 11 pt, Letter page with margins left 1.5 in, right 1 in, top 1 in, bottom 1.5 in,
// compatibility mode 15 unless the name starts with `m14-`. Three short paragraphs, then a fixed
// two-column table (tblInd 720, 3600-twip columns, 108-twip cell margins, single borders) whose
// second cell of the first row holds "First" and an anchor paragraph ("Anchor", the drawing,
// " after", or a longer text for the wrap cases), then a long paragraph. Every object is a red
// 36 pt square (54 pt tall for the grow cases) unless it is a chart, whose chart area is green.

const NS = [
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
	'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
	'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
].join(' ');
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const emu = (pt: number) => Math.round(pt * 12700);
const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** `relativeFrom:offsetPt` or `relativeFrom@align`. */
function position(axis: 'H' | 'V', spec: string): string {
	const aligned = /^(\w+)@(\w+)$/.exec(spec);
	if (aligned)
		return `<wp:position${axis} relativeFrom="${aligned[1]}"><wp:align>${aligned[2]}</wp:align></wp:position${axis}>`;
	const [from, offset] = spec.split(':');
	return `<wp:position${axis} relativeFrom="${from}"><wp:posOffset>${emu(Number(offset))}</wp:posOffset></wp:position${axis}>`;
}

function wrapXml(wrap: string): string {
	switch (wrap) {
		case 'square':
			return '<wp:wrapSquare wrapText="bothSides"/>';
		case 'largest':
			return '<wp:wrapSquare wrapText="largest"/>';
		case 'tight':
			return '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>';
		case 'topAndBottom':
			return '<wp:wrapTopAndBottom/>';
		default:
			return '<wp:wrapNone/>';
	}
}

interface AnchorOptions {
	widthPt?: number;
	heightPt?: number;
	chart?: boolean;
	behind?: boolean;
	/** The shape's fill, red by default. */
	color?: string;
	/** The second chart part (the demonstration has two charts). */
	secondChart?: boolean;
}

function anchor(
	id: number,
	h: string,
	v: string,
	wrap: string,
	inCell: 0 | 1,
	options: AnchorOptions = {},
): string {
	const cx = emu(options.widthPt ?? 36);
	const cy = emu(options.heightPt ?? 36);
	const graphic = options.chart
		? `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${options.secondChart ? 'rIdChart2' : 'rIdChart'}"/></a:graphicData></a:graphic>`
		: `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${options.color ?? 'FF0000'}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>`;
	const name = options.chart ? `Chart ${id}` : `Rectangle ${id}`;
	return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251659264 + id}" behindDoc="${options.behind ? 1 : 0}" locked="0" layoutInCell="${inCell}" allowOverlap="1"><wp:simplePos x="0" y="0"/>${position('H', h)}${position('V', v)}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrapXml(wrap)}<wp:docPr id="${id}" name="${name}"/><wp:cNvGraphicFramePr/>${graphic}</wp:anchor></w:drawing></w:r>`;
}

const WORDS =
	'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega';
/** The long paragraph after the table and the wrap cases' cell text. */
export const LONG_TEXT = Array.from({ length: 3 }, () => WORDS).join(' ');
const BORDER = 'w:val="single" w:sz="4" w:space="0" w:color="auto"';
const cell = (content: string) =>
	`<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
function table(anchorCell: string): string {
	return (
		`<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblInd w:w="720" w:type="dxa"/><w:tblBorders><w:top ${BORDER}/><w:left ${BORDER}/><w:bottom ${BORDER}/><w:right ${BORDER}/><w:insideH ${BORDER}/><w:insideV ${BORDER}/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>` +
		`<w:tr>${cell(para('Left'))}${cell(anchorCell)}</w:tr>` +
		`<w:tr>${cell(para('Row two'))}${cell(para('Last'))}</w:tr></w:tbl>`
	);
}
function anchorCell(drawing: string, text = ''): string {
	return text
		? `${para('First')}<w:p><w:r><w:t xml:space="preserve">Anchor </w:t></w:r>${drawing}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
		: `${para('First')}<w:p><w:r><w:t xml:space="preserve">Anchor</w:t></w:r>${drawing}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`;
}
const BEFORE = para('Before one') + para('Before two') + para('Before three');

/** One case: its body XML, compatibility mode and whether it holds the chart. */
export interface CellFloatCase {
	body: string;
	mode: number;
	chart?: true;
	/** A second copy of the chart part (with its style and colours). */
	secondChart?: true;
}

function cellCase(
	h: string,
	v: string,
	wrap: string,
	inCell: 0 | 1,
	text = '',
	options: AnchorOptions = {},
	mode = 15,
): CellFloatCase {
	return {
		body: BEFORE + table(anchorCell(anchor(1, h, v, wrap, inCell, options), text)) + para(LONG_TEXT),
		mode,
	};
}

const H_SPECS: Record<string, string> = {
	column: 'column:18',
	character: 'character:18',
	margin: 'margin:18',
	page: 'page:18',
	leftMargin: 'leftMargin:18',
	rightMargin: 'rightMargin:18',
	insideMargin: 'insideMargin:18',
	outsideMargin: 'outsideMargin:18',
	'column-center': 'column@center',
	'column-right': 'column@right',
	'margin-right': 'margin@right',
	'page-center': 'page@center',
};
const V_SPECS: Record<string, string> = {
	paragraph: 'paragraph:9',
	line: 'line:9',
	margin: 'margin:9',
	page: 'page:9',
	topMargin: 'topMargin:9',
	bottomMargin: 'bottomMargin:9',
	insideMargin: 'insideMargin:9',
	outsideMargin: 'outsideMargin:9',
	'margin-bottom': 'margin@bottom',
	'page-center': 'page@center',
	'line-top': 'line@top',
};

/** Every case by name. */
export function cellFloatCases(): Record<string, CellFloatCase> {
	const cases: Record<string, CellFloatCase> = {};
	for (const wrap of ['square', 'none']) {
		for (const [name, spec] of Object.entries(H_SPECS))
			cases[`h-${name}-${wrap}`] = cellCase(spec, 'paragraph:9', wrap, 1);
		for (const [name, spec] of Object.entries(V_SPECS))
			cases[`v-${name}-${wrap}`] = cellCase('column:18', spec, wrap, 1);
		// Mode 15 ignores layoutInCell="0".
		for (const [name, h, v] of [
			['h-page', 'page:18', 'paragraph:9'],
			['v-page', 'column:18', 'page:9'],
			['h-margin', 'margin:18', 'paragraph:9'],
			['v-margin', 'column:18', 'margin:9'],
		] as const)
			cases[`${name}-lic0-${wrap}`] = cellCase(h, v, wrap, 0);
		// Mode 14 honours it.
		for (const inCell of [0, 1] as const)
			for (const [name, h, v] of [
				['h-margin', 'margin:18', 'paragraph:9'],
				['h-page', 'page:18', 'paragraph:9'],
				['v-margin', 'column:18', 'margin:9'],
				['v-page', 'column:18', 'page:9'],
			] as const)
				cases[`m14-${name}-lic${inCell}-${wrap}`] = cellCase(h, v, wrap, inCell, '', {}, 14);
	}
	cases.behind = cellCase('column:18', 'paragraph:9', 'none', 1, '', { behind: true });
	// The character frame after a space, and in a body paragraph.
	cases['h-character-space'] = cellCase('character:18', 'paragraph:9', 'none', 1, ' tail');
	cases['h-character-body'] = {
		body: `${BEFORE}<w:p><w:r><w:t xml:space="preserve">Body anchor</w:t></w:r>${anchor(1, 'character:18', 'line:9', 'none', 1)}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>${para(LONG_TEXT)}`,
		mode: 15,
	};
	// Text wrapping around an object in the cell, and rows growing for it.
	for (const wrap of ['square', 'largest', 'tight', 'topAndBottom', 'none']) {
		cases[`wrap-${wrap}`] = cellCase('column:0', 'paragraph:0', wrap, 1, LONG_TEXT);
		cases[`wrap-right-${wrap}`] = cellCase('column@right', 'paragraph:0', wrap, 1, LONG_TEXT);
		cases[`grow-${wrap}`] = cellCase('column:18', 'paragraph:30', wrap, 1, '', { heightPt: 54 });
	}
	cases['grow-square-lic0'] = cellCase('column:18', 'paragraph:30', 'square', 0, '', {
		heightPt: 54,
	});
	// Floating charts: in the body (text wraps around them) and in a cell.
	const chart = { chart: true, widthPt: 180, heightPt: 120 } as const;
	cases['chart-body'] = {
		body: `${BEFORE}<w:p><w:r><w:t>Chart anchor</w:t></w:r>${anchor(1, 'margin:36', 'paragraph:6', 'square', 1, chart)}</w:p>${para(LONG_TEXT)}${para(LONG_TEXT)}`,
		mode: 15,
		chart: true,
	};
	cases['chart-body-topAndBottom'] = {
		body: `${BEFORE}<w:p><w:r><w:t>Chart anchor</w:t></w:r>${anchor(1, 'page@center', 'margin:100', 'topAndBottom', 1, chart)}</w:p>${para(LONG_TEXT)}${para(LONG_TEXT)}`,
		mode: 15,
		chart: true,
	};
	cases['chart-cell'] = {
		...cellCase('column:9', 'paragraph:6', 'square', 1, '', {
			chart: true,
			widthPt: 144,
			heightPt: 96,
		}),
		chart: true,
	};
	return cases;
}

/** The chart parts a chart case takes from `docx/__fixtures__/chart.docx`. */
export const CHART_PARTS = [
	'word/charts/chart1.xml',
	'word/charts/style1.xml',
	'word/charts/colors1.xml',
	'word/charts/_rels/chart1.xml.rels',
	'word/theme/theme1.xml',
] as const;

/**
 * A case's package parts (path to XML text). `chartParts` are `CHART_PARTS` read from chart.docx;
 * the chart area is filled green so the PDF script can find the chart's rectangle.
 */
export function cellFloatPackage(
	entry: CellFloatCase,
	chartParts: Record<string, string> = {},
): Record<string, string> {
	let overrides =
		'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>';
	let rels =
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';
	const files: Record<string, string> = {};
	if (entry.chart) {
		overrides +=
			'<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/word/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/><Override PartName="/word/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>';
		rels +=
			'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>';
		if (entry.secondChart) {
			overrides +=
				'<Override PartName="/word/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/word/charts/style2.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/><Override PartName="/word/charts/colors2.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>';
			rels +=
				'<Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart2.xml"/>';
		}
		for (const part of CHART_PARTS) {
			const xml = chartParts[part];
			if (xml === undefined) throw new Error(`A chart case needs ${part}`);
			files[part] =
				part === 'word/charts/chart1.xml'
					? xml.replace(
							'<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525"',
							'<c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:ln w="9525"',
						)
					: xml;
		}
		if (entry.secondChart) {
			files['word/charts/chart2.xml'] = chartParts['word/charts/chart1.xml']!;
			files['word/charts/style2.xml'] = chartParts['word/charts/style1.xml']!;
			files['word/charts/colors2.xml'] = chartParts['word/charts/colors1.xml']!;
			files['word/charts/_rels/chart2.xml.rels'] = chartParts['word/charts/_rels/chart1.xml.rels']!
				.replace('colors1.xml', 'colors2.xml')
				.replace('style1.xml', 'style2.xml');
		}
	}
	const sect =
		'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="2160" w:left="2160" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
	return {
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
		'_rels/.rels':
			'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
		'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
		'word/styles.xml': `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
		'word/settings.xml': `<?xml version="1.0" encoding="UTF-8"?><w:settings ${W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${entry.mode}"/></w:compat></w:settings>`,
		'word/document.xml': `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${entry.body}${sect}</w:body></w:document>`,
		...files,
	};
}

/**
 * The live demonstration: a square-wrapped shape, a chart, a top-and-bottom shape and a shape
 * behind the text, each in a table cell, then a chart floating in the body with text around it.
 */
export function cellFloatDemo(): CellFloatCase {
	const chart = { chart: true, widthPt: 144, heightPt: 96 } as const;
	const body =
		para('Floating objects in table cells') +
		table(
			anchorCell(
				anchor(1, 'column:0', 'paragraph:0', 'square', 1, { color: '4472C4' }),
				LONG_TEXT.slice(0, 220),
			),
		) +
		para('') +
		table(anchorCell(anchor(2, 'column:9', 'paragraph:6', 'square', 1, chart))) +
		para('') +
		table(
			anchorCell(
				anchor(3, 'column@center', 'paragraph:0', 'topAndBottom', 1, { color: 'ED7D31' }),
				LONG_TEXT.slice(0, 120),
			),
		) +
		para('') +
		table(
			anchorCell(
				anchor(4, 'column:24', 'paragraph:0', 'none', 1, {
					color: 'FFC000',
					behind: true,
					widthPt: 96,
				}),
				LONG_TEXT.slice(0, 120),
			),
		) +
		`<w:p><w:r><w:t>A floating chart in the body</w:t></w:r>${anchor(5, 'margin@right', 'paragraph:6', 'square', 1, { ...chart, widthPt: 180, heightPt: 120, secondChart: true })}</w:p>` +
		para(LONG_TEXT) +
		para(LONG_TEXT);
	return { body, mode: 15, chart: true, secondChart: true };
}
