import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { XMLParser } from 'fast-xml-parser';
import { expect, it } from 'vitest';
import { hasTextProperties, type PptxElement, type TextPptxElement } from '../../core/types';
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

it.each(['<a:buChar char="•"/>', '<a:buAutoNum type="arabicPeriod"/>'])(
	'retains list and soft-break XML while authoring inherited spacing: %s',
	async (bullet) => {
		const source = await readFile(new URL('../fixtures/e2e/linked-textbox.pptx', import.meta.url));
		const zip = await JSZip.loadAsync(source);
		const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const body = `<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:lvl1pPr></a:lstStyle><a:p><a:pPr><a:lnSpc><a:spcPct val="120000"/></a:lnSpc></a:pPr><a:r><a:t>First</a:t></a:r></a:p><a:p><a:pPr>${bullet}</a:pPr><a:r><a:t>Second</a:t></a:r><a:br/><a:r><a:t>continued</a:t></a:r></a:p><a:p><a:endParaRPr lang="fr-FR"/></a:p></p:txBody>`;
		zip.file('ppt/slides/slide1.xml', xml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/u, body));
		const handler = new PptxHandler();
		const reopened = new PptxHandler();
		try {
			const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
			const target = data.slides[0]!.elements.find(
				(e) => hasTextProperties(e) && e.text?.startsWith('First'),
			)!;
			if (!hasTextProperties(target)) throw new Error('Missing source shape');
			const update = withExclusiveLineSpacing({ lineSpacing: 1.2 });
			Object.assign(target, paragraphOverridePatch(target, update), {
				textStyle: mergeTextStyleUpdate(target.textStyle, update),
			});
			const bytes = await handler.save(data.slides);
			const savedZip = await JSZip.loadAsync(bytes);
			const savedXml = await savedZip.file('ppt/slides/slide1.xml')!.async('string');
			const savedBody = savedXml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/u)![0];
			const paragraphs = savedBody.match(/<a:p>[\s\S]*?<\/a:p>/gu)!;
			expect(paragraphs).toHaveLength(3);
			for (const paragraph of paragraphs) expect(paragraph).toContain('<a:spcPct val="120000"');
			expect(savedBody).toContain('val="150000"');
			expect(paragraphs[1]).toContain(
				bullet.startsWith('<a:buChar') ? '<a:buChar' : '<a:buAutoNum',
			);
			expect(paragraphs[1]).toContain('<a:br');
			expect(paragraphs[2]).toContain('lang="fr-FR"');
			const result = await reopened.load(bytes);
			const restored = result.slides[0]!.elements.find(
				(e) => hasTextProperties(e) && e.text?.startsWith('First'),
			)!;
			if (!hasTextProperties(restored)) throw new Error('Missing restored shape');
			expect(restored.textStyle?.inheritedParagraphSpacing?.[0]?.lineSpacing).toBe(1.5);
		} finally {
			handler.dispose();
			reopened.dispose();
		}
	},
);
