import type { ShapeStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import {
	coerceInspectorFieldValue,
	inspectorFieldPatch,
	inspectorTogglePatch,
	resolveInspectorSections,
} from './inspector-field-model';
import { SHAPE_EFFECT_SECTIONS } from './shape-effect-sections';

const patch = (s: ShapeStyle | undefined, section: string, field: string, raw: string | boolean) =>
	inspectorFieldPatch(SHAPE_EFFECT_SECTIONS, section, field, s, raw);

describe('shape effect sections', () => {
	it('lists the Format Shape > Effects sections in order, all off for a bare shape', () => {
		const views = resolveInspectorSections(SHAPE_EFFECT_SECTIONS, {});
		expect(views.map((v) => v.id)).toStrictEqual([
			'outerShadow',
			'innerShadow',
			'glow',
			'softEdge',
			'reflection',
			'blur',
			'bevel',
		]);
		expect(views.every((v) => v.toggleable && !v.on && v.fields.length === 0)).toBe(true);
	});

	it('uses only translation keys the English dictionary defines', () => {
		const keys = new Set<string>();
		for (const section of SHAPE_EFFECT_SECTIONS) {
			if (section.labelKey) keys.add(section.labelKey);
			for (const field of section.fields) {
				keys.add(field.labelKey);
				for (const option of field.options ?? []) keys.add(option.labelKey);
			}
		}
		const dictionary = translationsEn;
		expect([...keys].filter((key) => !(key in dictionary))).toStrictEqual([]);
	});

	it('gives every control a unique test id', () => {
		const ids = SHAPE_EFFECT_SECTIONS.flatMap((s) => [
			s.toggle?.testId,
			...s.fields.map((f) => f.testId),
		]);
		expect(ids.every(Boolean)).toBe(true);
		expect(new Set(ids).size).toBe(ids.length);
	});

	it('enables the outer shadow with its defaults and shows its fields', () => {
		const on = inspectorTogglePatch(SHAPE_EFFECT_SECTIONS, 'outerShadow', {}, true)!;
		expect(on).toMatchObject({ shadowColor: '#000000', shadowOpacity: 0.35, shadowAngle: 315 });
		const view = resolveInspectorSections(SHAPE_EFFECT_SECTIONS, on)[0]!;
		expect(view.on).toBe(true);
		expect(view.fields.map((f) => [f.id, f.control])).toStrictEqual([
			['color', 'color'],
			['opacity', 'range'],
			['blur', 'number'],
			['angle', 'number'],
			['distance', 'number'],
			['offsetX', 'number'],
			['offsetY', 'number'],
			['rotateWithShape', 'checkbox'],
		]);
		expect(view.fields[1]!.value).toBe(35);
		expect(inspectorTogglePatch(SHAPE_EFFECT_SECTIONS, 'outerShadow', on, false)).toStrictEqual({
			shadowColor: 'transparent',
		});
	});

	it('keeps the shadow offsets, angle and distance consistent', () => {
		const s: ShapeStyle = { shadowColor: '#000000', shadowAngle: 0, shadowDistance: 10 };
		const angled = patch(s, 'outerShadow', 'angle', '450')!;
		expect(angled.shadowAngle).toBe(90);
		expect(angled.shadowOffsetX).toBeCloseTo(0);
		expect(angled.shadowOffsetY).toBeCloseTo(10);

		const moved = patch(s, 'outerShadow', 'offsetY', '10')!;
		expect(moved.shadowAngle).toBeCloseTo(45);
		expect(moved.shadowDistance).toBeCloseTo(Math.sqrt(200));
	});

	it('writes a percent slider as a 0-1 opacity and clamps number boxes', () => {
		const s: ShapeStyle = { glowColor: '#ff0000', glowRadius: 4 };
		expect(patch(s, 'glow', 'opacity', '40')).toMatchObject({ glowOpacity: 0.4 });
		expect(patch(s, 'glow', 'radius', '500')).toMatchObject({ glowRadius: 96 });
		expect(patch(s, 'glow', 'radius', '')).toBeUndefined();
		expect(patch(s, 'glow', 'radius', 'abc')).toBeUndefined();
	});

	it('edits the reflection, soft edge, blur and bevel', () => {
		const reflection = inspectorTogglePatch(SHAPE_EFFECT_SECTIONS, 'reflection', {}, true)!;
		expect(reflection.reflectionStartOpacity).toBe(50);
		expect(patch(reflection, 'reflection', 'end', '20')).toBeUndefined();
		expect(patch(reflection, 'reflection', 'endOpacity', '20')).toMatchObject({
			reflectionEndOpacity: 20,
			reflectionStartOpacity: 50,
		});
		expect(inspectorTogglePatch(SHAPE_EFFECT_SECTIONS, 'softEdge', {}, true)).toStrictEqual({
			softEdgeRadius: 6,
		});
		expect(patch({ blurRadius: 4 }, 'blur', 'radius', '8')).toStrictEqual({ blurRadius: 8 });

		const bevel = inspectorTogglePatch(SHAPE_EFFECT_SECTIONS, 'bevel', {}, true)!;
		expect(bevel.shape3d).toMatchObject({ bevelTopType: 'circle', bevelTopWidth: 76200 });
		const bevelView = resolveInspectorSections(SHAPE_EFFECT_SECTIONS, bevel)[6]!;
		expect(bevelView.fields.find((f) => f.id === 'bevelTopWidth')!.value).toBe(6);
		expect(patch(bevel, 'bevel', 'bevelTopWidth', '3')!.shape3d!.bevelTopWidth).toBe(38100);
		expect(
			inspectorTogglePatch(SHAPE_EFFECT_SECTIONS, 'bevel', bevel, false)!.shape3d,
		).toStrictEqual({});
	});

	it('coerces toggles and wraps angles', () => {
		const angle = SHAPE_EFFECT_SECTIONS[0]!.fields.find((f) => f.id === 'angle')!;
		expect(coerceInspectorFieldValue(angle, '-90')).toBe(270);
		const rotate = SHAPE_EFFECT_SECTIONS[0]!.fields.find((f) => f.id === 'rotateWithShape')!;
		expect(coerceInspectorFieldValue(rotate, false)).toBe(false);
	});
});
