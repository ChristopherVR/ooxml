import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

import JSZip from 'jszip';

import { hasTextProperties, PptxHandler } from '../../../src/core/dist/pptx/index.mjs';
import { applyTextStyleUpdate } from '../../../src/ui/src/pptx/render/text-style-update.ts';

const output = process.argv[2];
if (!output) throw new Error('Usage: bun generate.ts <output-directory>');
await mkdir(output, { recursive: true });
const evidence = JSON.parse(await readFile(new URL('./evidence.json', import.meta.url), 'utf8'));
const zip = await JSZip.loadAsync(
	await readFile(
		new URL('../../../src/core/pptx/__tests__/fixtures/e2e/linked-textbox.pptx', import.meta.url),
	),
);
const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
zip.file(
	'ppt/slides/slide1.xml',
	xml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/gu, evidence.sourceTextBodyXml),
);
const source = await zip.generateAsync({ type: 'nodebuffer' });
await writeFile(resolve(output, 'source.pptx'), source);
const handler = new PptxHandler();
try {
	const data = await handler.load(source);
	const target = data.slides[0]!.elements.find(
		(e) => hasTextProperties(e) && e.text?.startsWith('First'),
	)!;
	Object.assign(target, applyTextStyleUpdate(target, { lineSpacing: 1.2 }, null).patch);
	await writeFile(
		resolve(output, 'core-edited.pptx'),
		new Uint8Array(await handler.save(data.slides)),
	);
} finally {
	handler.dispose();
}
