// Isolate explicit DrawingML paragraph boundaries, then verify them by reopening in Excel.
import JSZip from 'jszip';
import { join } from 'node:path';
import { NS, parseXml, buildXml, first, children } from '../src/core/xml';

const folder = process.argv[2];
if (!folder) throw new Error('Supply the native title reference folder');
const zip = await JSZip.loadAsync(await Bun.file(join(folder, 'mixed-lines.xlsx')).arrayBuffer());
const doc = parseXml(await zip.file('xl/charts/chart1.xml')!.async('string'));
const title = first(first(doc.documentElement, 'chart', NS.c), 'title', NS.c);
const rich = first(first(title, 'tx', NS.c), 'rich', NS.c)!;
const paragraph = first(rich, 'p', NS.a)!;
const next = doc.createElementNS(NS.a, 'a:p');
const properties = first(paragraph, 'pPr', NS.a);
if (properties) next.appendChild(properties.cloneNode(true));
for (const run of children(paragraph, 'r', NS.a)) {
	const content = first(run, 't', NS.a)?.textContent;
	if (content?.trim() === '' && content.includes('\n')) paragraph.removeChild(run);
	if (content === 'Forecast') next.appendChild(run);
}
rich.appendChild(next);
zip.file('xl/charts/chart1.xml', buildXml(doc));
await Bun.write(
	join(folder, 'mixed-paragraphs.xlsx'),
	await zip.generateAsync({ type: 'uint8array' }),
);
