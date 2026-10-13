import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { expect, it } from 'vitest';
import type { PptxElement, TextPptxElement } from '../../core/types';
import { PptxHandler } from '../../core/PptxHandler';
import { mergeTextStyleUpdate } from '../../editor/render/font-family-update';
import {
	paragraphOverridePatch,
	withExclusiveLineSpacing,
} from '../../editor/render/text-style-scope';

it('saves percent spacing over exact spacing while retaining paragraph levels and independent indents', async () => {
	const { handler, data } = await PptxHandler.createBlank({ initialSlideCount: 1 });
	const reopened = new PptxHandler();
	try {
		const element: TextPptxElement = {
			type: 'text',
			id: 'scope',
			x: 40,
			y: 40,
			width: 400,
			height: 200,
			text: 'First\nSecond',
			textStyle: { fontSize: 24, lineSpacingExactPt: 18 },
			paragraphIndents: [
				{ marginLeft: 12, indent: 5 },
				{ marginLeft: 24, indent: 7 },
			],
			textSegments: [
				{
					text: 'First',
					style: { fontSize: 24 },
					paragraphProperties: {
						lineSpacingExactPt: 24,
						paragraphMarginLeft: 12,
						paragraphIndent: 5,
					},
				},
				{ text: '\n', style: { fontSize: 24 }, isParagraphBreak: true },
				{
					text: 'Second',
					style: { fontSize: 24 },
					paragraphLevel: 2,
					paragraphProperties: { lineSpacing: 1.5, paragraphMarginLeft: 24, paragraphIndent: 7 },
				},
			],
		};
		const before = structuredClone(element);
		const update = withExclusiveLineSpacing({ lineSpacing: 2, paragraphMarginLeft: 36 });
		const patch = paragraphOverridePatch(element, update);
		expect(element).toEqual(before);
		Object.assign(element, patch, { textStyle: mergeTextStyleUpdate(element.textStyle, update) });
		data.slides[0]!.elements.push(element as PptxElement);
		const bytes = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(bytes);
		const slide = new XMLParser({ ignoreAttributes: false }).parse(
			await zip.file('ppt/slides/slide1.xml')!.async('string'),
		);
		const shape = slide['p:sld']['p:cSld']['p:spTree']['p:sp'];
		const paragraphs = shape['p:txBody']['a:p'];
		expect(paragraphs).toHaveLength(2);
		for (const paragraph of paragraphs) {
			expect(paragraph['a:pPr']['a:lnSpc']).toEqual({ 'a:spcPct': { '@_val': '200000' } });
		}
		const saved = await reopened.load(bytes);
		const target = saved.slides[0]!.elements.find(
			(entry) => entry.type === 'text' && entry.text === element.text,
		);
		if (!target || target.type !== 'text') throw new Error('Missing saved paragraphs');
		const first = target.textSegments?.find((segment) => segment.text === 'First');
		const second = target.textSegments?.find((segment) => segment.text === 'Second');
		expect(first?.paragraphProperties).toMatchObject({
			lineSpacing: 2,
			paragraphMarginLeft: 36,
			paragraphIndent: 5,
		});
		expect(second?.paragraphProperties).toMatchObject({
			lineSpacing: 2,
			paragraphMarginLeft: 36,
			paragraphIndent: 7,
		});
		expect(first?.paragraphProperties?.lineSpacingExactPt).toBeUndefined();
		expect(second?.paragraphLevel).toBe(2);
	} finally {
		handler.dispose();
		reopened.dispose();
	}
});
