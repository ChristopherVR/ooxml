import JSZip from 'jszip';
import { expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { captureNativeAnimationTemplate } from '../../core/services/animation-paint-template';
import { must } from '../../test-support/access';

async function targetTiming(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	const xml = await must(zip.file('ppt/slides/slide2.xml')).async('string');
	return must(xml.match(/<p:timing>[\s\S]*?<\/p:timing>/u)?.[0]);
}

it('writes painted behaviour once and keeps edits and node IDs across saves', async () => {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const sourceBuilder = createSlide('Blank').addText('Source', {
		x: 10,
		y: 10,
		width: 200,
		height: 40,
	});
	sourceBuilder.addAnimation(must(sourceBuilder.getLastElement()).id, {
		preset: 'flyIn',
		durationMs: 800,
	});
	data.slides.push(
		sourceBuilder.build(),
		createSlide('Blank')
			.addText('Target', {
				x: 10,
				y: 10,
				width: 200,
				height: 40,
			})
			.build(),
	);
	const original = await handler.save(data.slides);
	const editor = new PptxHandler();
	const deck = await editor.load(Uint8Array.from(original).buffer);
	const [source, target] = [must(deck.slides[0]), must(deck.slides[1])];
	const effect = must(source.nativeAnimations?.[0]);
	const template = must(
		captureNativeAnimationTemplate(must(source.rawTiming), must(effect.nodeId)),
	);
	const painted = {
		elementId: must(target.elements[0]).id,
		entrance: 'flyIn',
		durationMs: 1200,
		nativePaintTemplate: template,
	} as const;
	target.animations = [painted];
	target.isDirty = true;
	const first = await targetTiming(await editor.save(deck.slides));
	expect(first.match(/dur="1200"/gu)?.length).toBeGreaterThan(1);
	expect(target.animations[0]?.nativePaintTemplate).toBeUndefined();
	const ids = (timing: string) => timing.match(/<p:cTn id="\d+"/gu);
	target.animations = [{ ...must(target.animations[0]), durationMs: 1750 }];
	const retimed = await targetTiming(await editor.save(deck.slides));
	expect(retimed.match(/dur="1750"/gu)?.length).toBeGreaterThan(1);
	expect(ids(retimed)).toEqual(ids(first));
	target.animations = [{ ...must(target.animations[0]), direction: 'fromLeft' }];
	const redirected = await targetTiming(await editor.save(deck.slides));
	expect(redirected).toMatch(/presetSubtype="8"/u);
	expect(await targetTiming(await editor.save(deck.slides))).toBe(redirected);
});
