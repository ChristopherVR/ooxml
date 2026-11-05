import { describe, expect, it } from 'vitest';

import { exactDeviceBox } from './device-pixel-exact';
import { snappedStrokeTransform } from './device-pixel-stroke';

describe('exactDeviceBox', () => {
	// 60 pt by 50.5 pt at (340.3 pt, 290.4 pt), as the parser keeps it: whole
	// model pixels next to the exact EMU.
	const rect = {
		x: 454,
		y: 387,
		width: 80,
		height: 67,
		xEmu: 4321810,
		yEmu: 3688080,
		widthEmu: 762000,
		heightEmu: 641350,
	};

	it('positions an element at its EMU rounded to PowerPoint master units (1/576 inch)', () => {
		const box = exactDeviceBox(rect);
		// 3688080 EMU is 2323.2 master units: PowerPoint paints it at 2323.
		expect(box.y).toBeCloseTo(2323 / 6, 6);
		expect(box.height).toBeCloseTo(404 / 6, 6);
		expect(box.width).toBeCloseTo(80, 6);
	});

	it('keeps the model value where the EMU is missing or no longer matches it', () => {
		expect(exactDeviceBox({ x: 5, y: 6, width: 7, height: 8 })).toStrictEqual({
			x: 5,
			y: 6,
			width: 7,
			height: 8,
		});
		// Moved since load: the model's pixel no longer rounds to the stored EMU.
		expect(exactDeviceBox({ ...rect, x: 500 }).x).toBe(500);
		expect(exactDeviceBox({ ...rect, widthEmu: Number.NaN }).width).toBe(80);
	});

	it('moves the far edge of a snapped box to its own device pixel', () => {
		const model = snappedStrokeTransform('box', 454, 387, 1, { width: 80, height: 67 });
		const exact = snappedStrokeTransform(
			'box',
			454,
			387,
			1,
			{ width: 80, height: 67 },
			exactDeviceBox(rect),
		);
		expect(exact).not.toBe(model);
		// The scale spans the exact 67.3333 px, not the model's 67.
		expect(exact).toContain('scale(calc((80px');
		expect(exact).toContain('calc((67.3333px');
	});
});
