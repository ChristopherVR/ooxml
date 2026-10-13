import { describe, expect, it } from 'vitest';
import type { PptxSlide, TextPptxElement } from '../../index';
import {
	captureSlideClipboard,
	createSlidePaste,
	insertSlidePaste,
	captureSlideClipboardTemplates,
	applySlidePasteTemplates,
} from './slide-operations';

const slide = (id: string, name: string): PptxSlide => ({
	id,
	rId: id,
	slideNumber: 1,
	notes: name,
	elements: [
		{
			id: `text-${id}`,
			type: 'text',
			x: 0,
			y: 0,
			width: 200,
			height: 30,
			text: name,
			textSegments: [{ text: name, style: { bold: true } }],
		},
	],
});

describe('slide paste', () => {
	it('retains template artwork from copy time after its source cache was edited and removed', () => {
		const source = slide('a', 'A');
		const template = slide('template', 'Original artwork').elements[0]!;
		const mapping = { a: [template] };
		const captured = captureSlideClipboardTemplates(mapping, ['a']);
		(template as TextPptxElement).text = 'Changed artwork';
		const request = createSlidePaste(captureSlideClipboard([source], ['a']), 1, captured);
		const result = applySlidePasteTemplates({}, request);
		expect((result[request.copies[0]!.id]![0] as TextPptxElement).text).toBe('Original artwork');
		expect(result[request.copies[0]!.id]![0]).not.toBe(template);
	});

	it('remaps an animation targeting a nested group member to the cloned child', () => {
		const source = slide('a', 'A');
		const child = source.elements[0]!;
		source.elements = [
			{ id: 'group', type: 'group', x: 0, y: 0, width: 200, height: 100, children: [child] },
		];
		source.animations = [
			{ elementId: child.id, entrance: 'fadeIn', trigger: 'onClick', durationMs: 750 },
		];
		const copy = createSlidePaste(captureSlideClipboard([source], ['a']), 1).copies[0]!;
		const group = copy.elements[0]!;
		if (group.type !== 'group') throw new Error('Expected group');
		expect(copy.animations![0]!.elementId).toBe(group.children[0]!.id);
		expect(copy.animations![0]!.elementId).not.toBe(child.id);
	});
	it('inserts copied content at the destination in clipboard order with fresh identities', () => {
		const slides = [slide('a', 'A'), slide('b', 'B'), slide('c', 'C')];
		const clipboard = captureSlideClipboard(slides, ['a', 'c']);
		const request = createSlidePaste(clipboard, 2),
			result = insertSlidePaste(slides, request);
		expect(result.map((item) => item.notes)).toEqual(['A', 'B', 'A', 'C', 'C']);
		expect(result.map((item) => item.slideNumber)).toEqual([1, 2, 3, 4, 5]);
		expect(new Set(result.map((item) => item.id)).size).toBe(5);
		expect(result[2]!.elements[0]!.id).not.toBe(slides[0]!.elements[0]!.id);
		expect(result[2]!.sourceSlideId).toBe('a');
	});

	it('keeps clipboard content after source edits/deletion and supports repeated independent pastes', () => {
		const source = slide('ppt/slides/slide1.xml', 'Captured');
		const clipboard = captureSlideClipboard([source], [source.id]);
		(source.elements[0] as TextPptxElement).text = 'Changed';
		source.notes = 'Changed';
		const destination = [slide('b', 'B')];
		const first = insertSlidePaste(destination, createSlidePaste(clipboard, 1));
		const second = insertSlidePaste(first, createSlidePaste(clipboard, 2));
		expect(second.map((item) => item.notes)).toEqual(['B', 'Captured', 'Captured']);
		(second[1]!.elements[0] as TextPptxElement).textSegments![0]!.text = 'Edited copy';
		expect((second[2]!.elements[0] as TextPptxElement).textSegments![0]!.text).toBe('Captured');
		expect(second[2]!.sourceSlideId).toBe(source.id);
	});
});
