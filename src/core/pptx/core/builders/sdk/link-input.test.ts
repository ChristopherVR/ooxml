/**
 * Links on new elements and runs, including slide jumps by index, get their
 * relationship on save and read back as the same link.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import { createShapeElement, createTextElement } from './ElementFactory';
import type { SlideBuilder } from './SlideBuilder';

const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function save(build: (first: SlideBuilder) => void) {
	const { handler, createSlide } = await PptxHandler.create();
	const first = createSlide('Blank');
	build(first);
	const slides = [first.build()];
	slides.push(createSlide('Blank').addText('target').build());
	const zip = await JSZip.loadAsync(await handler.save(slides));
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
	const relTarget = (id: string) =>
		new RegExp(`<Relationship [^>]*Id="${id}"[^>]*>`).exec(rels)?.[0] ?? '';
	return { xml, rels, relTarget, bytes: await zip.generateAsync({ type: 'uint8array' }) };
}

const hlinks = (xml: string) => xml.match(/<a:hlinkClick\b[^>]*>/g) ?? [];
const rId = (node: string) => /r:id="([^"]*)"/.exec(node)?.[1] ?? '';

describe('link input', () => {
	it('writes a shape slide jump with its slide relationship', async () => {
		const { xml, relTarget } = await save((s) =>
			s.addShape('rect', { link: { slideIndex: 1, tooltip: 'Next' } }),
		);
		const [node] = hlinks(xml);
		expect(node).toContain('action="ppaction://hlinksldjump"');
		expect(node).toContain('tooltip="Next"');
		const rel = relTarget(rId(node));
		expect(rel).toContain('Target="slide2.xml"');
		expect(rel).toContain('relationships/slide"');
	});

	it('writes a run slide jump with its slide relationship', async () => {
		const { xml, relTarget } = await save((s) =>
			s.addText([
				{ text: 'go ', style: { bold: true } },
				{ text: 'there', style: { link: { slideIndex: 1 } } },
			]),
		);
		const [node] = hlinks(xml);
		expect(node).toContain('action="ppaction://hlinksldjump"');
		expect(rId(node)).not.toBe('');
		expect(relTarget(rId(node))).toContain('Target="slide2.xml"');
	});

	it('writes a run slide jump on a single-segment body', async () => {
		const { xml, relTarget } = await save((s) =>
			s.addText([{ text: 'all of it', style: { link: { slideIndex: 1 } } }]),
		);
		const [node] = hlinks(xml);
		expect(relTarget(rId(node))).toContain('Target="slide2.xml"');
	});

	it('writes URL links on runs, text boxes and pictures as external relationships', async () => {
		const { xml, relTarget } = await save((s) =>
			s
				.addText([{ text: 'site', style: { link: { url: 'https://example.com/a' } } }], {
					link: { url: 'https://example.com/b' },
				})
				.addImage(PNG, { link: { url: 'https://example.com/c' } }),
		);
		const targets = hlinks(xml).map((node) => relTarget(rId(node)));
		expect(targets).toHaveLength(3);
		for (const url of ['a', 'b', 'c']) {
			const rel = targets.find((t) => t.includes(`https://example.com/${url}`));
			expect(rel).toContain('TargetMode="External"');
		}
	});

	it('reads the jumps back as slide jumps to slide 2', async () => {
		const { bytes } = await save((s) =>
			s
				.addShape('rect', { link: { slideIndex: 1 } })
				.addText([{ text: 'x', style: { link: { slideIndex: 1 } } }]),
		);
		const data = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
		const [shape, text] = data.slides[0].elements;
		expect(shape.actionClick?.targetSlideIndex).toBe(1);
		expect(
			(text as { textSegments?: Array<{ style?: { hyperlinkTargetSlideIndex?: number } }> })
				.textSegments?.[0]?.style?.hyperlinkTargetSlideIndex,
		).toBe(1);
	});

	it('writes no tooltip when none is given', async () => {
		const { xml } = await save((s) => s.addShape('rect', { link: { slideIndex: 1 } }));
		expect(hlinks(xml)[0]).not.toContain('tooltip=');
	});

	it('writes a jump to the slide itself (index 0)', async () => {
		const { xml, relTarget } = await save((s) =>
			s
				.addShape('rect', { link: { slideIndex: 0 } })
				.addText([{ text: 'here', style: { link: { slideIndex: 0 } } }]),
		);
		const nodes = hlinks(xml);
		expect(nodes).toHaveLength(2);
		for (const node of nodes) {
			expect(relTarget(rId(node))).toContain('Target="slide1.xml"');
		}
	});

	it('leaves out and reports a slide jump past the last slide, so PowerPoint can open the file', async () => {
		const { handler, createSlide } = await PptxHandler.create();
		const slide = createSlide('Blank')
			.addShape('rect', { link: { slideIndex: 5, tooltip: 'gone' } })
			.addText([{ text: 'far', style: { link: { slideIndex: 7 } } }, { text: ' near' }])
			.addShape('ellipse', { link: { url: 'https://example.com/' } });
		const bytes = await handler.save([slide.build()]);
		const zip = await JSZip.loadAsync(bytes);
		const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
		const rels = (await zip.file('ppt/slides/_rels/slide1.xml.rels')?.async('string')) ?? '';
		// An empty relationship id on a slide jump makes PowerPoint refuse the file.
		expect(xml).not.toContain('r:id=""');
		expect(xml).not.toContain('hlinksldjump');
		expect(rels).not.toContain('relationships/slide"');
		// The other link on the slide is still written.
		expect(hlinks(xml)).toHaveLength(1);
		const missing = handler
			.getCompatibilityWarnings()
			.filter((w) => w.code === 'SAVE_SLIDE_JUMP_TARGET_MISSING');
		expect(missing.map((w) => w.message)).toEqual([
			expect.stringContaining('slide index 5'),
			expect.stringContaining('slide index 7'),
		]);
		const loaded = await new PptxHandler().load(bytes);
		const [shape, text, ellipse] = loaded.slides[0]?.elements ?? [];
		expect(shape?.actionClick).toBeUndefined();
		const segments = (
			text as { textSegments?: Array<{ text: string; style?: { hyperlink?: string } }> }
		).textSegments;
		expect(segments?.some((s) => s.style?.hyperlink)).toBe(false);
		expect(segments?.map((s) => s.text).join('')).toBe('far near');
		expect(ellipse?.actionClick?.url).toBe('https://example.com/');
	});

	it('keeps a reloaded slide jump on the same relationship when saved again', async () => {
		const { bytes } = await save((s) =>
			s
				.addShape('rect', { link: { slideIndex: 1, tooltip: 'Next' } })
				.addText([{ text: 'x', style: { link: { slideIndex: 1 } } }]),
		);
		const handler = new PptxHandler();
		const data = await handler.load(bytes.buffer as ArrayBuffer);
		const again = await JSZip.loadAsync(await handler.save(data.slides));
		const xml = (await again.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
		const rels = (await again.file('ppt/slides/_rels/slide1.xml.rels')?.async('string')) ?? '';
		const nodes = hlinks(xml);
		expect(nodes).toHaveLength(2);
		for (const node of nodes) {
			const rel = new RegExp(`<Relationship [^>]*Id="${rId(node)}"[^>]*>`, 'u').exec(rels)?.[0];
			expect(rel).toContain('Target="slide2.xml"');
		}
		expect(nodes[0]).toContain('tooltip="Next"');
		expect(handler.getCompatibilityWarnings().map((w) => w.code)).not.toContain(
			'SAVE_SLIDE_JUMP_TARGET_MISSING',
		);
	});

	it('rejects a slide index that is not a position', () => {
		expect(() => createShapeElement('rect', { link: { slideIndex: -1 } })).toThrow(RangeError);
		expect(() => createShapeElement('rect', { link: { slideIndex: 1.5 } })).toThrow(RangeError);
		expect(() =>
			createTextElement([{ text: 'x', style: { link: { slideIndex: Number.NaN } } }]),
		).toThrow(RangeError);
	});

	it('rejects an empty address', () => {
		expect(() => createShapeElement('rect', { link: { url: '' } })).toThrow(RangeError);
		expect(() => createShapeElement('rect', { link: { url: '  ' } })).toThrow(RangeError);
	});
});
