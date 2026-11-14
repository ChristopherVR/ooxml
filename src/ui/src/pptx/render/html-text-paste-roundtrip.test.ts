import JSZip from 'jszip';
import { PptxHandler } from 'ooxml-core/pptx';
import type { TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { htmlToPasteSegments, segmentsBodyText, spliceTextSegments } from './html-text-paste';

describe('pasted html survives save and reload', () => {
	it('writes links, bullets, numbering, highlight, baseline and margins PowerPoint can read', async () => {
		const { handler, data, createSlide } = await PptxHandler.createBlank();
		data.slides = [createSlide('Blank').addText('Lead').build()];
		const box = data.slides[0]!.elements.find((element) => element.type === 'text')!;
		const body = (box as { textSegments?: TextSegment[] }).textSegments ?? [
			{ text: 'Lead', style: {} },
		];
		const pasted = htmlToPasteSegments(
			'<p><a href="https://example.com/doc">link</a> <mark>hi</mark> x<sup>2</sup></p>' +
				'<ul><li>bullet</li></ul><ol><li>one</li><li>two</li></ol>',
			body[0]?.style ?? {},
		)!;
		const textSegments = spliceTextSegments(body, 4, 4, pasted.segments);
		Object.assign(box, { text: segmentsBodyText(textSegments), textSegments });

		const bytes = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(bytes);
		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
		const rId = /<a:hlinkClick[^>]*r:id="([^"]+)"/u.exec(slideXml)?.[1];
		expect(rId).toBeDefined();
		expect(rels).toMatch(new RegExp(`Id="${rId}"[^>]*Target="https://example.com/doc"`, 'u'));
		expect(slideXml).toContain('<a:buChar char="\u2022"');
		expect(slideXml).toMatch(/<a:buAutoNum type="arabicPeriod"/u);
		expect(slideXml).toContain('baseline="30000"');
		expect(slideXml).toMatch(/<a:highlight><a:srgbClr val="FFFF00"/u);

		const reloaded = new PptxHandler();
		const loaded = await reloaded.load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const text = loaded.slides[0]!.elements.find((element) => element.type === 'text') as {
			textSegments?: TextSegment[];
		};
		const segments = text.textSegments ?? [];
		const find = (needle: string) => segments.find((segment) => segment.text === needle);
		expect(find('link')?.style.hyperlink).toBe('https://example.com/doc');
		expect(find('hi')?.style.highlightColor?.toUpperCase()).toBe('#FFFF00');
		expect(find('2')?.style.baseline).toBe(30000);
		expect(segments.find((segment) => segment.bulletInfo?.char === '\u2022')).toBeDefined();
		expect(
			segments.filter((segment) => segment.bulletInfo?.autoNumType === 'arabicPeriod'),
		).toHaveLength(2);
		handler.dispose();
		reloaded.dispose();
	});
});
