import { describe, it, expect } from 'vitest';
import { PresentationBuilder } from '../builders/sdk/PresentationBuilder';
import { inspectPresentationText, setPresentationRunText } from './suite-text';

describe('suite presentation text operations', () => {
	it('replaces one run and preserves its neighbouring text', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		const slide = createSlide()
			.addText('Launch brief', { x: 40, y: 40, fontSize: 24 })
			.addText('Keep this text', { x: 40, y: 120 })
			.build();
		data.slides.push(slide);
		const bytes = await handler.save(data.slides);
		const before = await inspectPresentationText(bytes);
		const target = before.slides.find((s) => s.elements.some((e) => e.text === 'Launch brief'))!;
		const element = target.elements.find((e) => e.text === 'Launch brief')!;
		const saved = await setPresentationRunText(
			bytes,
			target.index,
			element.id,
			0,
			'Revised launch',
		);
		const after = await inspectPresentationText(saved);
		expect(JSON.stringify(after)).toContain('Revised launch');
		expect(JSON.stringify(after)).toContain('Keep this text');
		await expect(setPresentationRunText(bytes, -1, element.id, 0, 'Bad')).rejects.toThrow();
	});
});

describe('presentation outlines', () => {
	it('builds a cover slide and bulleted content slides', async () => {
		const { createPresentationFromOutline } = await import('./suite-text');
		const bytes = await createPresentationFromOutline([
			{ title: 'Quarterly review' },
			{ title: 'Highlights', bullets: ['Revenue up', 'Two launches'], notes: 'Keep it short' },
		]);
		const text = await inspectPresentationText(bytes);
		expect(text.slides).toHaveLength(2);
		const body = text.slides[1]!.elements.find((e) => e.text.includes('Revenue up'))!;
		expect(body.runs?.map((r) => r.text).filter((t) => !/^•?\s*$/.test(t))).toEqual([
			'Revenue up',
			'Two launches',
		]);
		const JSZip = (await import('jszip')).default;
		const zip = await JSZip.loadAsync(bytes);
		const slide = await zip.file('ppt/slides/slide2.xml')!.async('string');
		expect(slide.match(/<a:buChar/g)).toHaveLength(2);
		await expect(createPresentationFromOutline([])).rejects.toThrow();
	});
});
