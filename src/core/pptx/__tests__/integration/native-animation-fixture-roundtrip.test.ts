import { readFile } from 'node:fs/promises';

import JSZip from 'jszip';
import { expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import { editNativeEffectTiming } from '../../core/services/native-animation-edit';
import { must } from '../../test-support/access';

async function slideTiming(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	const xml = await must(zip.file('ppt/slides/slide1.xml')).async('string');
	return must(xml.match(/<p:timing>[\s\S]*?<\/p:timing>/u)?.[0]);
}

it('keeps untouched PowerPoint timing XML byte-identical after editing one effect', async () => {
	const bytes = await readFile(
		new URL('../fixtures/e2e/native-animation-timing.pptx', import.meta.url),
	);
	const handler = new PptxHandler();
	const deck = await handler.load(Uint8Array.from(bytes).buffer);
	const slide = must(deck.slides[0]);
	const animations = must(slide.nativeAnimations);
	// Use the same writer for the baseline because it expands empty XML tags.
	slide.isDirty = true;
	const original = await slideTiming(await handler.save(deck.slides));
	// The last click group is independent of the earlier After Previous chain.
	const nodeId = 17;
	Object.assign(
		slide,
		must(
			editNativeEffectTiming(must(slide.rawTiming), animations, nodeId, {
				field: 'duration',
				value: 1700,
			}),
		),
		{ isDirty: true },
	);
	const saved = await slideTiming(await handler.save(deck.slides));
	const editedEffect = /<p:cTn id="17" [\s\S]*?<\/p:childTnLst><\/p:cTn>/u;
	expect(original.match(editedEffect)?.[0]).toContain('dur="1000"');
	expect(saved.match(editedEffect)?.[0]).toContain('dur="1700"');
	expect(saved.replace(editedEffect, '')).toBe(original.replace(editedEffect, ''));
});
