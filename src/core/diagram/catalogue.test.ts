import { describe, expect, it } from 'vitest';
import {
	DIAGRAM_COLOR_ENTRIES,
	DIAGRAM_INSERT_CATEGORIES,
	DIAGRAM_INSERT_PRESETS,
	DIAGRAM_STYLE_ENTRIES,
	diagramColorLabelParams,
} from './catalogue';

describe('SmartArt gallery catalogue', () => {
	it('lists each layout once, under a known category, with starting text', () => {
		const layouts = DIAGRAM_INSERT_PRESETS.map((preset) => preset.layout);
		expect(new Set(layouts).size).toBe(layouts.length);
		const categories = new Set(DIAGRAM_INSERT_CATEGORIES.map((category) => category.id));
		for (const preset of DIAGRAM_INSERT_PRESETS) {
			expect(categories.has(preset.category), preset.layout).toBe(true);
			expect(preset.defaultItems.length, preset.layout).toBeGreaterThan(1);
		}
		for (const category of categories)
			expect(DIAGRAM_INSERT_PRESETS.some((preset) => preset.category === category)).toBe(true);
	});

	it('names the colour schemes and style intensities the model carries', () => {
		expect(DIAGRAM_COLOR_ENTRIES.map((entry) => entry.id)).toEqual([
			'colorful1',
			'colorful2',
			'colorful3',
			'monochromatic1',
			'monochromatic2',
		]);
		expect(DIAGRAM_STYLE_ENTRIES.map((entry) => entry.id)).toEqual(['flat', 'moderate', 'intense']);
		expect(diagramColorLabelParams(DIAGRAM_COLOR_ENTRIES[1]!)).toEqual({ from: 2, to: 3 });
		expect(diagramColorLabelParams(DIAGRAM_COLOR_ENTRIES[4]!)).toEqual({ accent: 2 });
		expect(diagramColorLabelParams(DIAGRAM_COLOR_ENTRIES[0]!)).toBeUndefined();
	});
});
