import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { PptxHandler } from '../../core/PptxHandler';
import { resolveOpacitySet } from '../../core/services/animation-opacity-set';
import { createOpacitySetDeck } from '../helpers/opacity-set-deck';

it.each([
	['indefinite duration', 'dur="1000"', 'dur="indefinite"'],
	['time filter', 'dur="1000"', 'dur="1000" tmFilter="0,0;1,1"'],
	['end sync attribute', 'dur="1000"', 'dur="1000" endSync="all"'],
	[
		'end sync child',
		'<p:childTnLst>',
		'<p:endSync evt="onEnd"><p:tn val="5"/></p:endSync><p:childTnLst>',
	],
	[
		'end condition',
		'<p:childTnLst>',
		'<p:endCondLst><p:cond evt="onClick" delay="0"/></p:endCondLst><p:childTnLst>',
	],
])(
	'refuses and preserves parent %s without changing ordinary click starts',
	async (_name, from, to) => {
		const zip = await JSZip.loadAsync(await createOpacitySetDeck(true));
		const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		// Change only the effect parent, not its child behavior or sequence containers.
		const changed = xml.replace(/(<p:cTn id="4"[\s\S]*?)(<p:set>)/u, (match) =>
			match.replace(from, to),
		);
		expect(changed).not.toBe(xml);
		zip.file('ppt/slides/slide1.xml', changed);
		const bytes = await zip.generateAsync({ type: 'nodebuffer' });
		const handler = new PptxHandler();
		try {
			const source = await handler.load(
				bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
			);
			expect(resolveOpacitySet(source.slides[0]!.nativeAnimations![0]!)).toBeUndefined();
			const saved = await handler.save(source.slides);
			const savedZip = await JSZip.loadAsync(saved);
			expect(await savedZip.file('ppt/slides/slide1.xml')!.async('string')).toContain(to);
			const reopened = await handler.load(
				saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength),
			);
			expect(resolveOpacitySet(reopened.slides[0]!.nativeAnimations![0]!)).toBeUndefined();
		} finally {
			handler.dispose();
		}
	},
);

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
