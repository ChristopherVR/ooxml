// Builds the variants `smartart-word-probe.ps1` opens in Word: a Word-authored SmartArt
// (`docx/__fixtures__/smartart-word.docx`) with its data part, cached drawing or definitions
// changed behind Word's back. Usage (from src/core): bun scripts/docx/smartart-word-variants.mjs <src.docx> <out folder>
import JSZip from 'jszip';
import { readFile, writeFile } from 'node:fs/promises';

const [, , src, outDir] = process.argv;
const base = await readFile(src);
const DOC = '{C5A4AE57-6EBF-4C2C-9489-D0E620B77143}';
const g = (n) => `{0000000${n}-AAAA-4BBB-8CCC-DDDDDDDDDDDD}`;

const addNode = (xml, text) => {
	const pt =
		`<dgm:pt modelId="${g(1)}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></dgm:t></dgm:pt>` +
		`<dgm:pt modelId="${g(2)}" type="parTrans" cxnId="${g(4)}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>` +
		`<dgm:pt modelId="${g(3)}" type="sibTrans" cxnId="${g(4)}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>`;
	const cxn = `<dgm:cxn modelId="${g(4)}" srcId="${DOC}" destId="${g(1)}" srcOrd="5" destOrd="0" parTransId="${g(2)}" sibTransId="${g(3)}"/>`;
	return xml
		.replace('</dgm:ptLst>', `${pt}</dgm:ptLst>`)
		.replace('</dgm:cxnLst>', `${cxn}</dgm:cxnLst>`);
};
const dropExt = (xml) => xml.replace(/<dgm:extLst>.*?<\/dgm:extLst>/s, '');
const dropDrawing = async (zip) => {
	zip.remove('word/diagrams/drawing1.xml');
	const rels = await zip.file('word/_rels/document.xml.rels').async('string');
	zip.file('word/_rels/document.xml.rels', rels.replace(/<Relationship Id="rId8"[^>]*\/>/, ''));
	const ct = await zip.file('[Content_Types].xml').async('string');
	zip.file(
		'[Content_Types].xml',
		ct.replace(/<Override PartName="\/word\/diagrams\/drawing1.xml"[^>]*\/>/, ''),
	);
};
// Only content points and parOf connections: what a fabricated or rewritten data part carries.
const dropPres = (xml) =>
	xml
		.replace(/<dgm:pt modelId="[^"]+" type="pres">.*?<\/dgm:pt>/gs, '')
		.replace(/<dgm:cxn [^>]*type="pres(?:Of|ParOf)"[^>]*\/>/g, '');

const variants = {
	'stale-text': async (zip, data) => {
		zip.file('word/diagrams/data1.xml', data.replace('>Beta<', '>Beta EDITED<'));
	},
	'stale-add': async (zip, data) => {
		zip.file('word/diagrams/data1.xml', addNode(data, 'Zeta'));
	},
	'nodrawing-add': async (zip, data) => {
		zip.file('word/diagrams/data1.xml', dropExt(addNode(data, 'Zeta')));
		await dropDrawing(zip);
	},
	'dangling-ext': async (zip, data) => {
		// The drawing part and rel stay, but the data part no longer names it.
		zip.file('word/diagrams/data1.xml', dropExt(addNode(data, 'Zeta')));
	},
	'nodrawing-nopres': async (zip, data) => {
		zip.file('word/diagrams/data1.xml', dropPres(dropExt(addNode(data, 'Zeta'))));
		await dropDrawing(zip);
	},
	'stale-nopres': async (zip, data) => {
		// Presentation points gone but the stale drawing (with its old pres ids) kept.
		zip.file('word/diagrams/data1.xml', dropPres(addNode(data, 'Zeta')));
	},
};

for (const [name, apply] of Object.entries(variants)) {
	const zip = await JSZip.loadAsync(base);
	const data = await zip.file('word/diagrams/data1.xml').async('string');
	await apply(zip, data);
	await writeFile(`${outDir}/${name}.docx`, await zip.generateAsync({ type: 'nodebuffer' }));
	console.log(name);
}

// The data unchanged, the cached drawing's text altered: does Word paint the cache?
{
	const zip = await JSZip.loadAsync(base);
	const drawing = await zip.file('word/diagrams/drawing1.xml').async('string');
	zip.file('word/diagrams/drawing1.xml', drawing.replace('>Gamma<', '>Cacheonly<'));
	await writeFile(`${outDir}/cache-only.docx`, await zip.generateAsync({ type: 'nodebuffer' }));
	console.log('cache-only');
}

// Stub definitions holding only a built-in uniqueId: does Word substitute its own copy?
{
	const NS =
		'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
	const stub = (el, id, extra = '') =>
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<dgm:${el} ${NS} uniqueId="${id}">${extra}</dgm:${el}>`;
	const zip = await JSZip.loadAsync(base);
	zip.file(
		'word/diagrams/colors1.xml',
		stub('colorsDef', 'urn:microsoft.com/office/officeart/2005/8/colors/colorful1'),
	);
	zip.file(
		'word/diagrams/quickStyle1.xml',
		stub('styleDef', 'urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1'),
	);
	zip.file(
		'word/diagrams/layout1.xml',
		stub(
			'layoutDef',
			'urn:microsoft.com/office/officeart/2005/8/layout/process1',
			'<dgm:layoutNode name="x"/>',
		),
	);
	const data = await zip.file('word/diagrams/data1.xml').async('string');
	zip.file(
		'word/diagrams/data1.xml',
		data
			.replace('layout/default"', 'layout/process1"')
			.replace('colors/accent1_2', 'colors/colorful1'),
	);
	await writeFile(
		`${outDir}/stub-definitions.docx`,
		await zip.generateAsync({ type: 'nodebuffer' }),
	);
	console.log('stub-definitions');
}
