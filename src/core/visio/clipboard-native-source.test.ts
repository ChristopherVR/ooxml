import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
	captureVisioClipboard,
	deserializeVisioClipboard,
	serializeVisioClipboard,
} from './clipboard';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { visioPasteCommand } from './ui/shape-clipboard';
import { visioDuplicateCommand } from './ui/shape-duplicate';

const directory = process.env['VISIO_NATIVE_DUPLICATE_DIR'];
/** Uses native source fixtures without reading or writing the operating-system clipboard. */
describe.skipIf(!directory)('native source clipboard import', () => {
	it('preserves native mixed XML and caches at three drawing scales for COM reopen', async () => {
		const bytes = new Uint8Array(await readFile(join(directory!, 'original.vsdx')));
		const source = await parseVsdx(bytes);
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			cases: { pageId: string; selection: string[] }[];
		};
		const pasteEdits: VisioEdit[] = [],
			duplicateEdits: VisioEdit[] = [];
		for (const item of evidence.cases) {
			const page = source.pages.find((candidate) => candidate.id === item.pageId)!;
			const clipboard = deserializeVisioClipboard(
				serializeVisioClipboard(await captureVisioClipboard(bytes, item.pageId, item.selection)),
			);
			pasteEdits.push(visioPasteCommand(page, clipboard)!);
			duplicateEdits.push(visioDuplicateCommand(page, item.selection)!);
		}
		const pasted = await editVsdx(bytes, pasteEdits),
			duplicated = await editVsdx(bytes, duplicateEdits);
		expect((await parseVsdx(pasted.bytes)).pages).toEqual(
			(await parseVsdx(duplicated.bytes)).pages,
		);
		await writeFile(join(directory!, 'clipboard-core.vsdx'), pasted.bytes);
	});
});
