// Builds src/core/visio/ui/org-chart-stencil-data.ts from Visio's own ORGCH_M.VSTX: a minimal
// stencil package holding the Executive, Manager, Position and Assistant Belt masters, the
// Dynamic connector and the style sheets they use. Provenance is recorded in PROVENANCE.md.
//   node scripts/gen-visio-org-stencil.mjs [path to ORGCH_M.VSTX]
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from '../src/core/node_modules/jszip/lib/index.js';

const here = dirname(fileURLToPath(import.meta.url));
const source =
	process.argv[2] ??
	'C:/Program Files/Microsoft Office/root/Office16/Visio Content/1033/ORGCH_M.VSTX';
const KEEP = [
	'Executive Belt',
	'Manager Belt',
	'Position Belt',
	'Assistant Belt',
	'Dynamic connector',
];

const zip = await JSZip.loadAsync(readFileSync(source));
const text = (path) => zip.file(path).async('string');
const mastersXml = await text('visio/masters/masters.xml');
const rels = await text('visio/masters/_rels/masters.xml.rels');
const documentXml = await text('visio/document.xml');

const out = new JSZip();
const entries = [...mastersXml.matchAll(/<Master ID='\d+'[\s\S]*?<\/Master>/g)].map((m) => m[0]);
const kept = [];
let index = 0;
let hasImage = false;
for (const entry of entries) {
	const name = /NameU='([^']*)'/.exec(entry)[1];
	if (!KEEP.includes(name)) continue;
	index++;
	const relId = /<Rel r:id='([^']*)'/.exec(entry)[1];
	const target = new RegExp(
		`Id="${relId}"[^>]*Target="([^"]*)"|Target="([^"]*)"[^>]*Id="${relId}"`,
	).exec(rels);
	const part = (target[1] ?? target[2]).replace(/^.*\//, '');
	let content = await text(`visio/masters/${part}`);
	// The Belt masters get their group text in front of the members from the Normal style of
	// Visio's own documents (DisplayMode 2); written on the root group, the position's text shows
	// in a drawing whose Normal style does not say so. The same value, so nothing else changes.
	if (/Type='Group'/.test(content) && !/N='DisplayMode'/.test(content.split('<Shape ID=')[1] ?? ''))
		content = content.replace(
			"<Cell N='ObjType' V='8'/>",
			"<Cell N='ObjType' V='8'/><Cell N='DisplayMode' V='2'/>",
		);
	out.file(`visio/masters/master${index}.xml`, content);
	const relsFile = zip.file(`visio/masters/_rels/${part}.rels`);
	if (relsFile) {
		out.file(`visio/masters/_rels/master${index}.xml.rels`, await relsFile.async('string'));
		hasImage = true;
	}
	kept.push(entry.replace(/<Rel r:id='[^']*'/, `<Rel r:id='rId${index}'`));
}
if (kept.length !== KEEP.length) throw new Error('A master is missing from the source.');
if (hasImage)
	out.file(
		'visio/media/image1.jpeg',
		await zip.file('visio/media/image1.jpeg').async('uint8array'),
	);

const head = /^[\s\S]*?<Masters[^>]*>/.exec(mastersXml)[0];
out.file('visio/masters/masters.xml', `${head}${kept.join('')}</Masters>`);
out.file(
	'visio/masters/_rels/masters.xml.rels',
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
		kept
			.map(
				(_, i) =>
					`<Relationship Id="rId${i + 1}" Type="http://schemas.microsoft.com/visio/2010/relationships/master" Target="master${i + 1}.xml"/>`,
			)
			.join('') +
		'</Relationships>',
);
const docHead = /^[\s\S]*?<VisioDocument[^>]*>/.exec(documentXml)[0];
const styles = /<StyleSheets>[\s\S]*?<\/StyleSheets>/.exec(documentXml)[0];
out.file('visio/document.xml', `${docHead}${styles}</VisioDocument>`);
out.file(
	'visio/_rels/document.xml.rels',
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/masters" Target="masters/masters.xml"/></Relationships>',
);
out.file(
	'_rels/.rels',
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/document" Target="visio/document.xml"/></Relationships>',
);
out.file(
	'[Content_Types].xml',
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.stencil.main+xml"/><Override PartName="/visio/masters/masters.xml" ContentType="application/vnd.ms-visio.masters+xml"/>' +
		kept
			.map(
				(_, i) =>
					`<Override PartName="/visio/masters/master${i + 1}.xml" ContentType="application/vnd.ms-visio.master+xml"/>`,
			)
			.join('') +
		'</Types>',
);

for (const name of Object.keys(out.files)) if (out.files[name].dir) delete out.files[name];
const bytes = await out.generateAsync({
	type: 'uint8array',
	compression: 'DEFLATE',
	compressionOptions: { level: 9 },
	// Fixed dates keep the output stable between runs.
	date: new Date(Date.UTC(2020, 0, 1)),
});
const b64 = Buffer.from(bytes).toString('base64');
const lines = b64.match(/.{1,200}/g).map((line) => `\t'${line}',`);
const ids = kept.map((entry) => /ID='(\d+)'/.exec(entry)[1]);
writeFileSync(
	join(here, '../src/core/visio/ui/org-chart-stencil-data.ts'),
	`// Generated by scripts/gen-visio-org-stencil.mjs from Visio's ORGCH_M.VSTX. Do not edit.
/** Master IDs inside the stencil: ${KEEP.map((n, i) => `${n} = ${ids[i]}`).join(', ')}. */
export const ORG_STENCIL_MASTER_IDS = {
${KEEP.map((n, i) => `\t'${n}': '${ids[i]}',`).join('\n')}
} as const;
const PARTS = [
${lines.join('\n')}
];
/** The stencil package bytes (a .vssx holding the org chart masters). */
export const orgStencilBytes = (): Uint8Array =>
	Uint8Array.from(atob(PARTS.join('')), (char) => char.charCodeAt(0));
`,
);
console.log(`stencil ${bytes.length} bytes, ids ${ids.join(',')}`);
