import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';

const FIXTURE = resolve(__dirname, '../fixtures/e2e/comment-mentions.pptx');
const buffer = (bytes: Uint8Array) =>
	bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

/** Load the deck, edit its mentioned comment's text, save and reload: what the mention covers then. */
async function editAndReload(followMentions: boolean): Promise<string[]> {
	const handler = new PptxHandler();
	const data = await handler.load(buffer(readFileSync(FIXTURE)));
	const comment = data.slides.flatMap((slide) => slide.comments ?? []).find((c) => c.mentions)!;
	expect(comment.text.slice(23, 33)).toBe('@Bob Grant');
	comment.text = `Hi! ${comment.text}`;
	if (followMentions) {
		comment.mentions = comment.mentions!.map((m) => ({ ...m, startIndex: m.startIndex + 4 }));
	}
	const reloaded = await new PptxHandler().load(buffer(await handler.save(data.slides)));
	const saved = reloaded.slides.flatMap((slide) => slide.comments ?? []).find((c) => c.mentions)!;
	return saved.mentions!.map((m) => saved.text.slice(m.startIndex, m.startIndex + m.length));
}

describe('modern comment mentions across an edit and a save', () => {
	it('saves the mention on its text whether or not the editor moved it', async () => {
		// The editor may leave the offsets as loaded (the save moves them) or move them with the edit.
		expect(await editAndReload(false)).toStrictEqual(['@Bob Grant']);
		expect(await editAndReload(true)).toStrictEqual(['@Bob Grant']);
	});
});
