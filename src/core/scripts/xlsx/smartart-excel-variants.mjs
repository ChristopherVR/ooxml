// Builds the variants `smartart-excel-probe.ps1` opens in Excel: an Excel-authored SmartArt
// (`xlsx/__fixtures__/excel-smartart.xlsx`, Basic Block List with Plan, Build, Ship) with its data
// part, cached drawing or definitions changed behind Excel's back. The Excel counterpart of
// `scripts/docx/smartart-word-variants.mjs`.
// Usage (from src/core): bun scripts/xlsx/smartart-excel-variants.mjs <src.xlsx> <out folder>
import JSZip from 'jszip';
import { readFile, writeFile } from 'node:fs/promises';

const [, , src, outDir] = process.argv;
const base = await readFile(src);
const DOC = '{3A2A5680-A619-4FD9-BD0B-8104EDB76D1A}';
const DATA = 'xl/diagrams/data1.xml';
const DRAWING = 'xl/diagrams/drawing1.xml';
const RELS = 'xl/drawings/_rels/drawing1.xml.rels';
const g = (n) => `{0000000${n}-AAAA-4BBB-8CCC-DDDDDDDDDDDD}`;

const addNode = (xml, text) => {
	const pt =
		`<dgm:pt modelId="${g(1)}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></dgm:t></dgm:pt>` +
		`<dgm:pt modelId="${g(2)}" type="parTrans" cxnId="${g(4)}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>` +
		`<dgm:pt modelId="${g(3)}" type="sibTrans" cxnId="${g(4)}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>`;
	const cxn = `<dgm:cxn modelId="${g(4)}" srcId="${DOC}" destId="${g(1)}" srcOrd="3" destOrd="0" parTransId="${g(2)}" sibTransId="${g(3)}"/>`;
	return xml
		.replace('</dgm:ptLst>', `${pt}</dgm:ptLst>`)
		.replace('</dgm:cxnLst>', `${cxn}</dgm:cxnLst>`);
};
const dropExt = (xml) => xml.replace(/<dgm:extLst>.*?<\/dgm:extLst>/s, '');
const dropDrawing = async (zip) => {
	zip.remove(DRAWING);
	const rels = await zip.file(RELS).async('string');
	zip.file(RELS, rels.replace(/<Relationship Id="rId5"[^>]*\/>/, ''));
	const ct = await zip.file('[Content_Types].xml').async('string');
	zip.file(
		'[Content_Types].xml',
		ct.replace(/<Override PartName="\/xl\/diagrams\/drawing1.xml"[^>]*\/>/, ''),
	);
};
const dropPres = (xml) =>
	xml
		.replace(/<dgm:pt modelId="[^"]+" type="pres">.*?<\/dgm:pt>/gs, '')
		.replace(/<dgm:cxn [^>]*type="pres(?:Of|ParOf)"[^>]*\/>/g, '');

const variants = {
	'stale-text': async (zip, data) => {
		zip.file(DATA, data.replace('>Build<', '>Build EDITED<'));
	},
	'stale-add': async (zip, data) => {
		zip.file(DATA, addNode(data, 'Zeta'));
	},
	'nodrawing-add': async (zip, data) => {
		zip.file(DATA, dropExt(addNode(data, 'Zeta')));
		await dropDrawing(zip);
	},
	'dangling-ext': async (zip, data) => {
		zip.file(DATA, dropExt(addNode(data, 'Zeta')));
	},
	'nodrawing-nopres': async (zip, data) => {
		zip.file(DATA, dropPres(dropExt(addNode(data, 'Zeta'))));
		await dropDrawing(zip);
	},
	'stale-nopres': async (zip, data) => {
		zip.file(DATA, dropPres(addNode(data, 'Zeta')));
	},
};

for (const [name, apply] of Object.entries(variants)) {
	const zip = await JSZip.loadAsync(base);
	await apply(zip, await zip.file(DATA).async('string'));
	await writeFile(`${outDir}/${name}.xlsx`, await zip.generateAsync({ type: 'nodebuffer' }));
	console.log(name);
}

// The data unchanged, the cached drawing's text altered: does Excel paint the cache?
{
	const zip = await JSZip.loadAsync(base);
	const drawing = await zip.file(DRAWING).async('string');
	zip.file(DRAWING, drawing.replace('>Ship<', '>Cacheonly<'));
	await writeFile(`${outDir}/cache-only.xlsx`, await zip.generateAsync({ type: 'nodebuffer' }));
	console.log('cache-only');
}
