import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { ELEMENT_GEOMETRY_SECTIONS, sizePositionSections } from './element-geometry-fields';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';

const frame = { x: 53.27, y: 10.5, width: 200.4, height: 99.6, rotation: 15.2 };

describe('element geometry section', () => {
	it('shows the frame rounded to whole px, untitled, in X Y W H rotation order', () => {
		const [view] = resolveInspectorSections(ELEMENT_GEOMETRY_SECTIONS, frame);
		expect(view!.labelKey).toBeUndefined();
		expect(view!.fields.map((f) => [f.id, f.value])).toStrictEqual([
			['x', 53],
			['y', 11],
			['width', 200],
			['height', 100],
			['rotation', 15],
		]);
	});

	it('labels the fields with dictionary keys ("X" is the accessible name)', () => {
		const keys = ELEMENT_GEOMETRY_SECTIONS[0]!.fields.map((f) => f.labelKey);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
		expect(translationsEn[keys[0]!]).toBe('X');
	});

	it('exposes return, transparency and duration settings for inserted slide and section zooms', () => {
		for (const zoomType of ['slide', 'section'] as const) {
			const zoom = { ...frame, type: 'zoom' as const, zoomType };
			const sections = sizePositionSections(zoom);
			const options = resolveInspectorSections(sections, zoom).find(
				({ id }) => id === 'zoomOptions',
			);
			expect(options?.fields.map(({ id }) => id)).toEqual([
				'zoomReturn',
				'zoomTransparent',
				'zoomDuration',
			]);
			expect(inspectorFieldPatch(sections, 'zoomOptions', 'zoomReturn', zoom, false)).toEqual({
				returnToParent: false,
			});
			expect(inspectorFieldPatch(sections, 'zoomOptions', 'zoomTransparent', zoom, true)).toEqual({
				zoomShowBackground: false,
			});
			expect(inspectorFieldPatch(sections, 'zoomOptions', 'zoomDuration', zoom, '750')).toEqual({
				transitionDurationMs: 750,
			});
		}
		expect(sizePositionSections({ ...frame, type: 'shape' })).toHaveLength(1);
	});

	it('changes one named Summary Zoom tile independently', () => {
		const first = { ...frame, sectionId: 'intro', title: 'Introduction', targetSlideIndex: 1 };
		const second = { ...frame, sectionId: 'details', title: 'Details', targetSlideIndex: 3 };
		const zoom = {
			...frame,
			type: 'zoom' as const,
			zoomType: 'summary' as const,
			summaryTargets: [first, second],
		};
		const sections = sizePositionSections(zoom);
		expect(
			resolveInspectorSections(sections, zoom)
				.slice(1)
				.map((section) => section.label),
		).toEqual(['Introduction', 'Details']);
		const patch = inspectorFieldPatch(sections, 'zoomTile-1', 'zoomDuration', zoom, '750');
		expect(patch?.summaryTargets).toEqual([first, { ...second, transitionDurationMs: 750 }]);
		expect(patch?.summaryTargets?.[0]).toBe(first);
		expect(zoom.summaryTargets[1]).toBe(second);
	});

	it('writes a geometry patch, keeping the size at least 1px', () => {
		const patch = (field: string, raw: string) =>
			inspectorFieldPatch(ELEMENT_GEOMETRY_SECTIONS, 'geometry', field, frame, raw);
		expect(patch('x', '222')).toStrictEqual({ x: 222 });
		expect(patch('width', '0')).toStrictEqual({ width: 1 });
		expect(patch('rotation', '-30')).toStrictEqual({ rotation: -30 });
		expect(patch('y', '')).toBeUndefined();
		expect(
			resolveInspectorSections(ELEMENT_GEOMETRY_SECTIONS, { x: 0, y: 0, width: 1, height: 1 })[0]!
				.fields[4]!.value,
		).toBe(0);
	});
});
