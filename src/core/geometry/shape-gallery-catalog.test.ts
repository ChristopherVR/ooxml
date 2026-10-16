import { describe, expect, it } from 'vitest';
import { PRESET_SHAPE_DEFINITIONS } from './preset-shape-definitions';
import { evaluatePresetShape } from './preset-shape-evaluator';
import {
	SHAPE_GALLERY_CATEGORIES,
	SHAPE_GALLERY_ENTRIES,
	shapeGalleryName,
} from './shape-gallery-catalog';

describe('shape gallery catalogue', () => {
	it('lists PowerPoint categories in order', () => {
		expect(SHAPE_GALLERY_CATEGORIES.map((category) => category.name)).toEqual([
			'Lines',
			'Rectangles',
			'Basic Shapes',
			'Block Arrows',
			'Equation Shapes',
			'Flowchart',
			'Stars and Banners',
			'Callouts',
			'Action Buttons',
		]);
	});

	it('draws every entry from the preset geometry tables', () => {
		for (const entry of SHAPE_GALLERY_ENTRIES) {
			const geometry = evaluatePresetShape(entry.preset, 40, 40);
			expect(geometry?.paths.length, entry.preset).toBeGreaterThan(0);
		}
	});

	it('places every preset the renderer defines exactly once', () => {
		const placed = SHAPE_GALLERY_ENTRIES.map((entry) => entry.preset);
		expect(new Set(placed).size).toBe(placed.length);
		const missing = PRESET_SHAPE_DEFINITIONS.map((definition) => definition.name).filter(
			(name) => !placed.includes(name),
		);
		expect(missing).toEqual([]);
	});

	it('names each entry', () => {
		expect(shapeGalleryName('star5')).toBe('Star: 5 Points');
		expect(shapeGalleryName('nope')).toBeUndefined();
	});
});
