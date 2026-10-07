// Test-reference preparation only. Product chart parsing/writing stays in the shared core.
import JSZip from 'jszip';
import { resolve, join } from 'node:path';
import { NS, parseXml, buildXml, elements, first } from '../src/core/xml';
import { builtInChartStyleXml } from '../src/core/chart/built-in-text-style';

const folder = resolve(process.argv[2] ?? '');
if (!process.argv[2]) throw new Error('Supply the output folder of record-xlsx-chart-styles.ps1');
for (const id of [
	...Array.from({ length: 48 }, (_, i) => i + 1),
	...Array.from({ length: 48 }, (_, i) => i + 101),
]) {
	const legacy = id > 100 ? id - 100 : id;
	const zip = await JSZip.loadAsync(
		await Bun.file(join(folder, `column-${legacy}.xlsx`)).arrayBuffer(),
	);
	const doc = parseXml(await zip.file('xl/charts/chart1.xml')!.async('string'));
	for (const node of elements(doc.documentElement))
		if (node.localName === 'AlternateContent' || node.localName === 'style')
			node.parentNode!.removeChild(node);
	for (const node of Array.from(doc.getElementsByTagNameNS(NS.c, 'txPr')))
		node.parentNode?.removeChild(node);
	// Keep empty rich-text property nodes: deleting them changes Excel's rich-text interpretation.
	for (const name of ['rPr', 'defRPr', 'endParaRPr'])
		for (const node of Array.from(doc.getElementsByTagNameNS(NS.a, name))) {
			for (const attr of Array.from(node.attributes))
				if (attr.localName !== 'lang') node.removeAttribute(attr.name);
			for (const child of Array.from(node.childNodes)) node.removeChild(child);
		}
	const fragment = parseXml(
		`<c:chartSpace xmlns:c="${NS.c}">${builtInChartStyleXml(id)}</c:chartSpace>`,
	);
	doc.documentElement.insertBefore(
		doc.importNode(fragment.documentElement.firstChild!, true),
		first(doc.documentElement, 'chart', NS.c) ?? null,
	);
	zip.file('xl/charts/chart1.xml', buildXml(doc));
	const rels = parseXml(await zip.file('xl/charts/_rels/chart1.xml.rels')!.async('string'));
	for (const node of elements(rels.documentElement))
		if (node.getAttribute('Type')?.endsWith('/chartStyle')) node.parentNode!.removeChild(node);
	zip.file('xl/charts/_rels/chart1.xml.rels', buildXml(rels));
	zip.remove('xl/charts/style1.xml');
	const types = parseXml(await zip.file('[Content_Types].xml')!.async('string'));
	for (const node of elements(types.documentElement))
		if (node.getAttribute('PartName') === '/xl/charts/style1.xml')
			node.parentNode!.removeChild(node);
	zip.file('[Content_Types].xml', buildXml(types));
	await Bun.write(
		join(folder, `builtin-${id}.xlsx`),
		await zip.generateAsync({ type: 'uint8array' }),
	);
}
