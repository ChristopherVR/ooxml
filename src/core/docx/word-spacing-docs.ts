// The documents of `scripts/docx-word-spacing-probe.ps1`, rebuilt in TypeScript so the layout can be
// compared with what Word measured (`__fixtures__/word-spacing/spacing-probe.json`). Keep the two
// in step: every group, label and spacing value here is the script's.
import JSZip from 'jszip';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const tw = (pt: number) => Math.trunc(pt * 20);
const para = (label: string, spacing = '', extra = '', run = '') =>
	`<w:p><w:pPr>${spacing ? `<w:spacing ${spacing}/>` : ''}${extra}</w:pPr><w:r><w:t>${label}</w:t></w:r>${run}</w:p>`;
const sect = (type: string) =>
	`<w:sectPr><w:type w:val="${type}"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
const exact = 'w:line="240" w:lineRule="exact"';
const LEVELS = [0, 6, 12, 18];
const THREE = [0, 12, 18];

function table(id: string, pa: number, cb: number, ca: number, pb: number): string {
	const cell = para(`${id}c`, `w:before="${tw(cb)}" w:after="${tw(ca)}"`);
	return (
		para(`${id}p`, `w:after="${tw(pa)}"`) +
		`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${cell}</w:tc></w:tr></w:tbl>` +
		para(`${id}q`, `w:before="${tw(pb)}"`)
	);
}

/** Group name to the body XML the probe script writes for it. */
export function probeBodies(): Record<string, { body: string; sect?: string }> {
	const out: Record<string, { body: string; sect?: string }> = {};
	let b = '';
	for (const a of LEVELS)
		for (const c of LEVELS)
			b += para(`a${a}b${c}x`, `w:after="${tw(a)}"`) + para(`a${a}b${c}y`, `w:before="${tw(c)}"`);
	out.grid = { body: b };
	b = para('first18', `w:before="${tw(18)}"`) + para('sep0');
	const auto: Record<string, [string, string]> = {
		autoA_autoB: ['w:after="100" w:afterAutospacing="1"', 'w:before="100" w:beforeAutospacing="1"'],
		autoA_b0: ['w:after="100" w:afterAutospacing="1"', ''],
		a0_autoB: ['', 'w:before="100" w:beforeAutospacing="1"'],
		autoA_b6: ['w:after="100" w:afterAutospacing="1"', `w:before="${tw(6)}"`],
		autoA_b18: ['w:after="100" w:afterAutospacing="1"', `w:before="${tw(18)}"`],
		a6_autoB: [`w:after="${tw(6)}"`, 'w:before="100" w:beforeAutospacing="1"'],
		a18_autoB: [`w:after="${tw(18)}"`, 'w:before="100" w:beforeAutospacing="1"'],
		a0_b0: ['', ''],
	};
	for (const [k, [x, y]] of Object.entries(auto)) b += para(`${k}x`, x) + para(`${k}y`, y);
	out.auto = { body: b };
	b = '';
	for (const pa of [0, 12])
		for (const cb of [0, 12])
			for (const ca of [0, 12])
				for (const pb of [0, 12]) b += table(`pa${pa}_cb${cb}_ca${ca}_pb${pb}`, pa, cb, ca, pb);
	out.table = { body: b };
	for (const [kind, type] of [
		['secnext', 'nextPage'],
		['seccont', 'continuous'],
	] as const) {
		b = '';
		for (const a of THREE)
			for (const c of THREE)
				b +=
					para(`a${a}b${c}x`, `w:after="${tw(a)}"`, sect(type)) +
					para(`a${a}b${c}y`, `w:before="${tw(c)}"`, sect(type));
		out[kind] = { body: b };
	}
	b = '';
	for (const a of THREE)
		for (const c of THREE)
			b +=
				para(`a${a}b${c}x`, `w:after="${tw(a)}"`, '', '<w:r><w:br w:type="page"/></w:r>') +
				para(`a${a}b${c}y`, `w:before="${tw(c)}"`);
	out.pgbrk = { body: b };
	b = para('first', exact);
	for (let i = 1; i < 54; i++) b += para(`f${i}`, exact);
	out.natural = { body: b + para('target18', `w:before="${tw(18)}"`) + para('target0') };
	b = '';
	for (const n of [1, 2, 3])
		for (const i of [1, 2, 3])
			b += para(
				`list${n}i${i}`,
				'',
				`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${n}"/></w:numPr>`,
			);
	out.list = { body: b };
	b = '';
	for (const a of THREE)
		for (const c of THREE)
			b +=
				para(`a${a}b${c}x`, `w:after="${tw(a)}"`) +
				`<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="${tw(c)}"/></w:pPr><w:r><w:t>a${a}b${c}y</w:t></w:r></w:p>`;
	out.pgbefore = { body: b };
	b = '';
	for (let i = 1; i <= 54; i++) b += para(`f${i}`, exact);
	out.cols = {
		body: b + para('target18', `w:before="${tw(18)}"`) + para('target0'),
		sect: `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="2" w:space="720"/></w:sectPr>`,
	};
	return out;
}

const bullet = (id: number, text: string, font: string) =>
	`<w:abstractNum w:abstractNumId="${id}"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/></w:rPr></w:lvl></w:abstractNum>`;

/** A `.docx` for one probe group in the given compatibility mode and HTML auto spacing option. */
export async function probeDocx(group: string, mode: number, noHtml: boolean): Promise<Uint8Array> {
	const { body, sect: bodySect } = probeBodies()[group]!;
	const ct = (name: string, type: string) =>
		`<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${ct('styles', 'styles')}${ct('numbering', 'numbering')}${ct('settings', 'settings')}</Types>`,
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	const rel = (id: number, type: string, target: string) =>
		`<Relationship Id="rId${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel(1, 'styles', 'styles.xml')}${rel(2, 'settings', 'settings.xml')}${rel(3, 'numbering', 'numbering.xml')}</Relationships>`,
	);
	zip.file(
		'word/styles.xml',
		`<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
	);
	zip.file(
		'word/numbering.xml',
		`<?xml version="1.0" encoding="UTF-8"?><w:numbering ${W}>${bullet(1, '&#xF0B7;', 'Symbol')}${bullet(2, 'o', 'Courier New')}<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num></w:numbering>`,
	);
	zip.file(
		'word/settings.xml',
		`<?xml version="1.0" encoding="UTF-8"?><w:settings ${W}><w:compat>${noHtml ? '<w:doNotUseHTMLParagraphAutoSpacing/>' : ''}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`,
	);
	zip.file(
		'word/document.xml',
		`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}${bodySect ?? sect('nextPage')}</w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
