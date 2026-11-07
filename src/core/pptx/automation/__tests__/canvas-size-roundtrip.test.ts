import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../index';
import { pptxTools } from '../catalogue/index';
import { createTestPptxBytes } from './helpers/create-test-pptx';

function canvasTool() {
	const tool = pptxTools.find((candidate) => candidate.name === 'set_canvas_size');
	if (tool?.kind !== 'edit') throw new Error('Missing canvas size tool');
	return tool;
}

describe('set_canvas_size round trip', () => {
	it('writes the requested dimensions and clears an obsolete size preset', async () => {
		const input = await createTestPptxBytes(1);
		const result = await canvasTool().run(input, { width: 720, height: 480.5 });
		const { bytes, ...report } = result;
		expect(report).toEqual({ canvasWidth: 720, canvasHeight: 480.5 });
		if (!bytes) throw new Error('The canvas edit did not save');

		const handler = new PptxHandler();
		try {
			const saved = await handler.load(new Uint8Array(bytes).buffer);
			expect(saved).toMatchObject({
				widthEmu: 9144000,
				heightEmu: 6102350,
				width: 960,
				height: 641,
			});
			expect(saved.slideSizeType).toBeUndefined();
			expect(saved.slides[0].elements).toHaveLength(2);
		} finally {
			handler.dispose();
		}
	});

	it.each([
		{ width: 71, height: 480 },
		{ width: 720, height: 4033 },
	])('rejects an out-of-range slide size ($width x $height)', async (size) => {
		await expect(canvasTool().run(await createTestPptxBytes(1), size)).rejects.toThrow(
			'Canvas width and height must be between 72 and 4032 points.',
		);
	});
});
