import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import type { PptxElement } from '../../core/types/elements';

/**
 * PowerPoint ignores `a:linkedTxbx` (it is Word's linked-text-box construct):
 * the head keeps all its text and successors stay empty. The loader still
 * parses the link and the writer must keep it, so the file is not degraded.
 */
const FIXTURE = resolve(__dirname, '../fixtures/e2e/linked-textbox.pptx');

function links(elements: PptxElement[]): [string, number | undefined, number | undefined][] {
	const out: [string, number | undefined, number | undefined][] = [];
	for (const el of elements) {
		if ('linkedTxbxId' in el && el.linkedTxbxId !== undefined) {
			out.push([el.id, Number(el.linkedTxbxId), el.linkedTxbxSeq]);
		}
		if (el.type === 'group') out.push(...links(el.children ?? []));
	}
	return out;
}

describe('a:linkedTxbx round trip', () => {
	it('keeps every chain link through save and reload', async () => {
		const bytes = readFileSync(FIXTURE);
		const first = await new PptxHandler().load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const before = links(first.slides[0]!.elements);
		expect(before.length).toBe(8);
		const handler = new PptxHandler();
		const loaded = await handler.load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const saved = await handler.save(loaded.slides);
		const again = await new PptxHandler().load(
			saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
		);
		expect(links(again.slides[0]!.elements).map(([, id, seq]) => [id, seq])).toEqual(
			before.map(([, id, seq]) => [id, seq]),
		);
	});
});
