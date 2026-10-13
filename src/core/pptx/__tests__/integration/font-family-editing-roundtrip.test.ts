import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { expect, it } from 'vitest';
import { PptxHandler } from '../../core/PptxHandler';
import { textRunsStylePatch } from '../../editor/render/text-runs-style-patch';

it('saves concrete picked faces without replaced theme tokens or unrelated timing edits', async () => {
	const source = await readFile(
		new URL('../fixtures/e2e/animation-after-effect.pptx', import.meta.url),
	);
	const original = await JSZip.loadAsync(source);
	const parser = new XMLParser({ ignoreAttributes: false, parseAttributeValue: false });
	const timing = parser.parse(await original.file('ppt/slides/slide1.xml')!.async('string'))[
		'p:sld'
	]['p:timing'];
	const handler = new PptxHandler();
	const reopened = new PptxHandler();
	try {
		const data = await handler.load(
			source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength),
		);
		const shape = data.slides[0]!.elements.find(
			(element) => element.type === 'shape' && element.text === 'RectDimRGB',
		);
		if (!shape || shape.type !== 'shape') throw new Error('Missing native text shape');
		shape.text = 'RectDimRGB Tail';
		shape.textStyle = { ...shape.textStyle, fontFamily: 'Old Sans', latinFontThemeToken: '+mn-lt' };
		shape.textSegments = [
			{
				text: 'RectDimRGB ',
				style: {
					fontFamily: 'Old Sans',
					eastAsiaFont: 'Old Sans',
					latinFontThemeToken: '+mn-lt',
					eastAsiaFontThemeToken: '+mn-ea',
					complexScriptFont: 'Arial',
					symbolFont: 'Wingdings',
					bold: true,
				},
			},
			{ text: 'Tail', style: { fontFamily: 'Old Sans', italic: true } },
		];
		Object.assign(shape, textRunsStylePatch(shape, { fontFamily: 'Georgia' }));
		const output = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(output);
		// Retained tree equality allows the existing writer's self-closing XML normalization.
		expect(
			parser.parse(await zip.file('ppt/slides/slide1.xml')!.async('string'))['p:sld']['p:timing'],
		).toEqual(timing);
		const saved = await reopened.load(output);
		const target = saved.slides[0]!.elements.find(
			(element) => element.type === 'shape' && element.text === 'RectDimRGB Tail',
		);
		if (!target || target.type !== 'shape') throw new Error('Missing saved text shape');
		expect(target.textSegments?.[0]?.style).toMatchObject({
			fontFamily: 'Georgia',
			eastAsiaFont: 'Georgia',
			complexScriptFont: 'Arial',
			symbolFont: 'Wingdings',
			bold: true,
		});
		expect(target.textSegments?.[0]?.style.latinFontThemeToken).toBeUndefined();
		expect(target.textSegments?.[0]?.style.eastAsiaFontThemeToken).toBeUndefined();
	} finally {
		handler.dispose();
		reopened.dispose();
	}
});
