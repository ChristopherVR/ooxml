import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

import JSZip from 'jszip';

import {
	PptxHandler,
	createSlideClipboardPackage,
	importSlidePackage,
} from '../../../src/core/dist/pptx/index.mjs';
import { renameFixtureParts } from '../../../src/core/pptx/__tests__/helpers/rename-package-parts';
import { createSlideClipboardDeck } from '../../../src/core/pptx/__tests__/helpers/slide-clipboard-deck';
import { buildXml, NS, parseXml } from '../../../src/core/xml/index';

const directory = process.argv[2];
if (!directory) throw new Error('Pass an output directory for synthetic acceptance artifacts');
const output = resolve(directory);
for (const variant of ['standard', 'boundary-names']) {
	const folder = resolve(output, variant);
	await mkdir(folder, { recursive: true });
	let source = new Uint8Array(await createSlideClipboardDeck(true));
	let destination = new Uint8Array(await createSlideClipboardDeck(false));
	if (variant === 'boundary-names') {
		source = await renameFixtureParts(
			source,
			new Map([
				['ppt/slideMasters/slideMaster1.xml', 'ppt/templates/company-design.xml'],
				['ppt/notesMasters/notesMaster1.xml', 'ppt/templates/presenter-notes.xml'],
			]),
		);
		const zip = await JSZip.loadAsync(destination);
		const presentation = parseXml(await zip.file('ppt/presentation.xml')!.async('string'));
		presentation.getElementsByTagNameNS(NS.p, 'sldId')[0]!.setAttribute('id', '2147483647');
		presentation.getElementsByTagNameNS(NS.p, 'sldMasterId')[0]!.setAttribute('id', '4294967295');
		zip.file('ppt/presentation.xml', buildXml(presentation));
		const master = parseXml(await zip.file('ppt/slideMasters/slideMaster1.xml')!.async('string'));
		master.getElementsByTagNameNS(NS.p, 'sldLayoutId')[0]!.setAttribute('id', '4294967294');
		zip.file('ppt/slideMasters/slideMaster1.xml', buildXml(master));
		destination = await zip.generateAsync({ type: 'uint8array' });
	}
	const clipboard = await createSlideClipboardPackage(source, [0]);
	const imported = await importSlidePackage(destination, clipboard, { insertAt: 1 });
	const handler = new PptxHandler();
	try {
		const data = await handler.load(
			imported.bytes.buffer.slice(
				imported.bytes.byteOffset,
				imported.bytes.byteOffset + imported.bytes.byteLength,
			) as ArrayBuffer,
		);
		const saved = await handler.save(data.slides);
		for (const [name, bytes] of [
			['source', source],
			['destination', destination],
			['clipboard', clipboard],
			['imported', imported.bytes],
			['core-resaved', saved],
		] as const)
			await writeFile(resolve(folder, `${name}.pptx`), bytes);
	} finally {
		handler.dispose();
	}
}
console.log(`Generated synthetic PowerPoint acceptance artifacts: ${output}`);
