// Test-reference preparation only; product parsing/writing uses the shared core.
import JSZip from 'jszip';
import { join, resolve } from 'node:path';
import { NS, parseXml, buildXml, first } from '../src/core/xml';

const folder = resolve(process.argv[2] ?? '');
if (!process.argv[2]) throw new Error('Supply the native chart reference folder');
for (const id of [2, 41, 102, 141, 201, 212]) {
	for (const mode of ['full', 'size', 'face', 'direct'] as const) {
		const source =
			id > 148
				? join(resolve(process.argv[3] ?? folder), `column-${id}.xlsx`)
				: join(folder, `builtin-${id}.xlsx`);
		const zip = await JSZip.loadAsync(await Bun.file(source).arrayBuffer());
		const doc = parseXml(await zip.file('xl/charts/chart1.xml')!.async('string'));
		const oldText = first(doc.documentElement, 'txPr', NS.c);
		if (oldText) doc.documentElement.removeChild(oldText);
		if (id > 148 && mode !== 'direct') {
			for (const node of Array.from(doc.getElementsByTagNameNS(NS.c, 'txPr')))
				node.parentNode?.removeChild(node);
			for (const name of ['rPr', 'defRPr', 'endParaRPr'])
				for (const node of Array.from(doc.getElementsByTagNameNS(NS.a, name))) {
					for (const attr of Array.from(node.attributes))
						if (attr.localName !== 'lang') node.removeAttribute(attr.name);
					for (const child of Array.from(node.childNodes)) node.removeChild(child);
				}
		}
		const properties =
			mode === 'size' ? 'sz="2000"' : mode === 'face' ? '' : 'sz="2000" b="0" i="1"';
		const children =
			mode === 'size'
				? ''
				: '<a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Arial"/>';
		const fragment = parseXml(
			`<c:txPr xmlns:c="${NS.c}" xmlns:a="${NS.a}"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr ${properties}>${children}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`,
		);
		doc.documentElement.insertBefore(
			doc.importNode(fragment.documentElement, true),
			first(doc.documentElement, 'printSettings', NS.c) ?? null,
		);
		if (mode === 'direct') {
			const chart = first(doc.documentElement, 'chart', NS.c);
			const title = first(chart, 'title', NS.c)!;
			const run = title.getElementsByTagNameNS(NS.a, 'rPr')[0]!;
			run.setAttribute('sz', '3200');
			run.setAttribute('b', '1');
		}
		zip.file('xl/charts/chart1.xml', buildXml(doc));
		await Bun.write(
			join(folder, `root-${id}-${mode}.xlsx`),
			await zip.generateAsync({ type: 'uint8array' }),
		);
	}
}
