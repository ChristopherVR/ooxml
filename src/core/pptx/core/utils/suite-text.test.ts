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
