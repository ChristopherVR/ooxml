import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { PptxHandler } from '../../core/PptxHandler';
import { resolveOpacitySet } from '../../core/services/animation-opacity-set';
import { createOpacitySetDeck } from '../helpers/opacity-set-deck';

it('preserves refused background and indefinite child timing profiles through save/reopen', async () => {
	const zip = await JSZip.loadAsync(await createOpacitySetDeck(false));
	const part = zip.file('ppt/slides/slide1.xml')!;
	const xml = (await part.async('string'))
		.replace(/<p:spTgt spid="(\d+)"\/>/gu, '<p:spTgt spid="$1"><p:bg/></p:spTgt>')
		.replace('id="5" dur="1000"', 'id="5" dur="1000" repeatCount="indefinite"');
	zip.file('ppt/slides/slide1.xml', xml);
	const bytes = await zip.generateAsync({ type: 'nodebuffer' });
	const handler = new PptxHandler();
	try {
		const source = await handler.load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
		);
		expect(resolveOpacitySet(source.slides[0]!.nativeAnimations![0]!)).toBeUndefined();
		const saved = await handler.save(source.slides);
		const savedZip = await JSZip.loadAsync(saved);
		const savedXml = await savedZip.file('ppt/slides/slide1.xml')!.async('string');
		expect(savedXml).toContain('repeatCount="indefinite"');
		expect(savedXml).toContain('<p:bg');
		const reopened = await handler.load(
			saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength),
		);
		expect(
			reopened.slides[0]!.nativeAnimations![0]!.behaviors?.[0]?.timing.unsupportedSetTiming,
		).toBeTruthy();
		expect(resolveOpacitySet(reopened.slides[0]!.nativeAnimations![0]!)).toBeUndefined();
	} finally {
		handler.dispose();
	}
});

it.each([true, false])(
	'preserves whole-target numeric string and child fill on save/reopen (hold=%s)',
	async (hold) => {
		const handler = new PptxHandler();
		try {
			const bytes = await createOpacitySetDeck(hold);
			const source = await handler.load(
				bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
			);
			const animation = source.slides[0]!.nativeAnimations![0]!;
			expect(animation.behaviors?.[0]?.timing.fill).toBe(hold ? 'hold' : 'remove');
			expect(resolveOpacitySet(animation)?.value).toBe(0.4);
			expect(resolveOpacitySet(animation)?.holdEndState).toBe(hold);
			const saved = await handler.save(source.slides);
			const reopened = await handler.load(
				saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength),
			);
			expect(resolveOpacitySet(reopened.slides[0]!.nativeAnimations![0]!)).toEqual(
				resolveOpacitySet(animation),
			);
			const zip = await JSZip.loadAsync(saved);
			const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
			expect(xml).toContain('style.opacity');
			expect(xml).toContain('val="0.4"');
		} finally {
			handler.dispose();
		}
	},
);
