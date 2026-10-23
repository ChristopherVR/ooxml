import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { ELEMENT_GEOMETRY_SECTIONS } from './element-geometry-fields';
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
