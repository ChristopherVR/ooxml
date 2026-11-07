/**
 * Rich text segments of a new text box save as given: a style every segment
 * shares is not lost or replaced by the element style. Each case is saved and
 * loaded back, so the written runs and the reloaded segments agree.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { TextSegment } from '../../types/text';
import type { SlideBuilder } from './SlideBuilder';

async function roundTrip(
	build: (slide: SlideBuilder) => void,
): Promise<{ runs: string[]; paragraphs: number; segments: TextSegment[] }> {
	const { handler, createSlide } = await PptxHandler.create();
	const slide = createSlide('Blank');
	build(slide);
	const bytes = await handler.save([slide.build()]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const loaded = await new PptxHandler().load(bytes);
	const element = loaded.slides[0]?.elements[0] as { textSegments?: TextSegment[] } | undefined;
	const segments = (element?.textSegments ?? []).filter((s) => !s.isParagraphBreak);
	return {
		runs: xml.match(/<a:r>[\s\S]*?<\/a:r>/gu) ?? [],
		paragraphs: xml.match(/<a:p>/gu)?.length ?? 0,
		segments,
	};
}

describe('rich text segment style', () => {
	it('keeps the style of a single segment', async () => {
		const { runs, segments } = await roundTrip((s) =>
			s.addText([{ text: 'A', style: { bold: true, italic: true, fontFamily: 'Roboto' } }]),
		);
		expect(runs).toHaveLength(1);
		expect(runs[0]).toContain('b="1"');
		expect(runs[0]).toContain('i="1"');
		expect(runs[0]).toContain('<a:latin typeface="Roboto"');
		expect(segments[0]?.style?.bold).toBe(true);
		expect(segments[0]?.style?.italic).toBe(true);
		expect(segments[0]?.style?.fontFamily).toBe('Roboto');
	});

	it('keeps a style that several segments share', async () => {
		const { runs, segments } = await roundTrip((s) =>
			s.addText([
				{ text: 'A', style: { color: '#FF0000', fontSize: 30 } },
				{ text: 'B', style: { color: '#FF0000', fontSize: 30 } },
			]),
		);
		for (const run of runs) {
			expect(run).toContain('<a:srgbClr val="FF0000"');
			expect(run).toContain('sz="2250"');
		}
		expect(segments.map((s) => s.style?.color?.toUpperCase())).toEqual(['#FF0000']);
	});

	it('does not let the element style replace a value every segment overrides', async () => {
		const { runs, segments } = await roundTrip((s) =>
			s.addText(
				[
					{ text: 'A', style: { bold: false, italic: true } },
					{ text: 'B', style: { bold: false } },
				],
				{ bold: true },
			),
		);
		for (const run of runs) {
			expect(run).not.toContain('b="1"');
		}
		expect(segments.some((s) => s.style?.bold === true)).toBe(false);
		expect(segments[0]?.style?.italic).toBe(true);
	});

	it('still applies an element option that no segment overrides', async () => {
		const { runs } = await roundTrip((s) =>
			s.addText([{ text: 'A' }, { text: 'B', style: { italic: true } }], { bold: true }),
		);
		expect(runs).toHaveLength(2);
		for (const run of runs) {
			expect(run).toContain('b="1"');
		}
	});

	it('keeps segments with different styles apart', async () => {
		const { runs } = await roundTrip((s) =>
			s.addText([
				{ text: 'A', style: { bold: true } },
				{ text: 'B', style: { strikethrough: true } },
			]),
		);
		expect(runs[0]).toContain('b="1"');
		expect(runs[0]).not.toContain('strike="sngStrike"');
		expect(runs[1]).toContain('strike="sngStrike"');
		expect(runs[1]).not.toContain('b="1"');
	});

	it('keeps a shared style across a newline inside a segment', async () => {
		const { runs, paragraphs } = await roundTrip((s) =>
			s.addText([{ text: 'one\ntwo', style: { bold: true } }]),
		);
		expect(paragraphs).toBe(2);
		expect(runs).toHaveLength(2);
		for (const run of runs) {
			expect(run).toContain('b="1"');
		}
	});

	it('writes a plain string as before', async () => {
		const { runs } = await roundTrip((s) => s.addText('Plain', { bold: true, fontSize: 20 }));
		expect(runs).toHaveLength(1);
		expect(runs[0]).toContain('b="1"');
		expect(runs[0]).toContain('sz="1500"');
	});
});
