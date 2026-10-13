import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createRichFindDeck } from '../helpers/find-replace-deck';
import { PptxHandler } from '../../core/PptxHandler';
import { findInSlides, replaceInSlides } from '../../editor/render/find-replace';
import { collectElementText } from '../../editor/render/slide-search';
import {
	captureSlideClipboard,
	createSlidePaste,
	insertSlidePaste,
	createBlankSlide,
} from '../../editor/render/slide-operations';

describe('find/replace saved OOXML', () => {
	it('saves a retained clipboard after its source was deleted, keeping grouped text, table formatting and notes', async () => {
		const handler = new PptxHandler(),
			data = await handler.load(await createRichFindDeck());
		data.slides[0]!.notes = 'Captured speaker notes';
		data.slides[0]!.hidden = true;
		const sourceGroup = data.slides[0]!.elements.find((element) => element.type === 'group');
		if (sourceGroup?.type !== 'group') throw new Error('Expected source group');
		data.slides[0]!.animations = [
			{
				elementId: sourceGroup.children[0]!.id,
				entrance: 'fadeIn',
				durationMs: 750,
				trigger: 'onClick',
			},
		];
		const clipboard = captureSlideClipboard(data.slides, [data.slides[0]!.id]);
		const next = insertSlidePaste([createBlankSlide(1)], createSlidePaste(clipboard, 1));
		const saved = await handler.save(next);
		const reopened = await new PptxHandler().load(
			saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
		);
		expect(reopened.slides).toHaveLength(2);
		expect(reopened.slides[1]!.notes).toContain('Captured speaker notes');
		expect(reopened.slides[1]!.hidden).toBe(true);
		const pastedGroup = reopened.slides[1]!.elements.find((element) => element.type === 'group');
		if (pastedGroup?.type !== 'group') throw new Error('Expected pasted group');
		expect(reopened.slides[1]!.animations).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					elementId: pastedGroup.children[0]!.id,
					entrance: 'fadeIn',
					durationMs: 750,
				}),
			]),
		);
		expect(findInSlides(reopened.slides.slice(1), 'Annual report')).toHaveLength(3);
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file(reopened.slides[1]!.id)!.async('string')).toContain('opaque-cell');
	});

	it('saves replacements in rich shapes, grouped text and table cells without flattening formats', async () => {
		const handler = new PptxHandler(),
			data = await handler.load(await createRichFindDeck());
		expect(findInSlides(data.slides, 'Annual report')).toHaveLength(3);
		const replaced = replaceInSlides(data.slides, 'report', 'summary');
		expect(replaced.replacements).toBe(3);
		const saved = await handler.save([...replaced.slides]);
		const zip = await JSZip.loadAsync(saved),
			xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		expect(xml).toContain('opaque-cell');
		expect(xml).toContain('rIdKeep');
		expect(xml).toContain('lang="fr-FR"');
		expect(xml).toContain('u="sng"');
		const reopened = await new PptxHandler().load(
			saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
		);
		expect(findInSlides(reopened.slides, 'Annual summary')).toHaveLength(3);
		expect(findInSlides(reopened.slides, 'report')).toHaveLength(0);
		for (const element of reopened.slides[0]!.elements) {
			if (element.type === 'text' || element.type === 'shape')
				expect(element.textSegments?.some((run) => run.style.bold)).toBe(true);
			if (element.type === 'table') {
				expect(element.tableData!.rows[0]!.cells[0]!.textRuns?.some((run) => run.italic)).toBe(
					true,
				);
				expect(collectElementText(element)).toContain('Annual summary');
			}
		}
	});
});

it('replaces rich speaker notes and reopens with body formatting preserved', async () => {
	const handler = new PptxHandler();
	const data = await handler.load(await createRichFindDeck());
	data.slides[0]!.notes = 'Speaker report';
	data.slides[0]!.notesSegments = [
		{ text: 'Speaker re', style: { bold: true } },
		{ text: 'port', style: { italic: true } },
	];
	const replaced = replaceInSlides(data.slides, 'report', 'summary', { includeNotes: true });
	expect(replaced.replacements).toBe(4);
	const saved = await handler.save([...replaced.slides]);
	const reopened = await new PptxHandler().load(
		saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
	);
	expect(reopened.slides[0]!.notes).toContain('Speaker summary');
	expect(
		reopened.slides[0]!.notesSegments?.find((run) => run.text.includes('summary'))?.style.bold,
	).toBe(true);
});
