import type { ShapeStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import {
	SHAPE_FILL_MODE_SECTIONS,
	SHAPE_FILL_OPTIONS_SECTIONS,
	shapeFillModeChange,
	shapeFillModeOf,
} from './shape-fill-fields';

const fieldIds = (style: ShapeStyle | undefined) =>
	resolveInspectorSections(SHAPE_FILL_OPTIONS_SECTIONS, style)[0]!.fields.map((f) => f.id);

describe('shape fill fields', () => {
	it('reads the fill kind the way PowerPoint names it', () => {
		expect(shapeFillModeOf(undefined)).toBe('solid');
		expect(shapeFillModeOf({ fillColor: 'transparent' })).toBe('none');
		expect(shapeFillModeOf({ fillMode: 'gradient' })).toBe('gradient');
		expect(shapeFillModeOf({ fillMode: 'image' })).toBe('image');
		expect(shapeFillModeOf({ fillMode: 'theme', fillColor: '#123456' })).toBe('solid');
		expect(shapeFillModeOf({ fillMode: 'image', useBackgroundFill: true })).toBe('background');
	});

	it('offers the six fill kinds with dictionary labels', () => {
		const [view] = resolveInspectorSections(SHAPE_FILL_MODE_SECTIONS, { fillMode: 'pattern' });
		const [field] = view!.fields;
		expect(field!.value).toBe('pattern');
		expect(field!.options!.map((o) => o.value)).toStrictEqual([
			'none',
			'solid',
			'gradient',
			'image',
			'pattern',
			'background',
		]);
		const keys = [...SHAPE_FILL_MODE_SECTIONS, ...SHAPE_FILL_OPTIONS_SECTIONS].flatMap((section) =>
			section.fields.flatMap((f) => [f.labelKey, ...(f.options ?? []).map((o) => o.labelKey)]),
		);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});

	it('shows each kind its own fields', () => {
		expect(fieldIds({ fillMode: 'none' })).toStrictEqual([]);
		expect(fieldIds({ fillMode: 'solid' })).toStrictEqual(['fillOpacity']);
		expect(fieldIds({ fillMode: 'gradient' })).toStrictEqual([
			'fillOpacity',
			'fillGradientType',
			'fillGradientAngle',
		]);
		expect(fieldIds({ fillMode: 'gradient', fillGradientType: 'radial' })).toStrictEqual([
			'fillOpacity',
			'fillGradientType',
		]);
		expect(fieldIds({ fillMode: 'image' })).toStrictEqual(['fillOpacity', 'fillImageMode']);
		expect(fieldIds({ useBackgroundFill: true })).toStrictEqual([]);
	});

	it('switches kinds with the defaults each needs', () => {
		const style: ShapeStyle = { fillColor: '#ff0000' };
		expect(shapeFillModeChange(style, 'none')).toStrictEqual({
			fillMode: 'none',
			fillColor: 'transparent',
		});
		expect(shapeFillModeChange({ fillColor: 'transparent' }, 'solid')).toStrictEqual({
			fillColor: '#3b82f6',
			fillMode: 'solid',
		});
		expect(shapeFillModeChange(style, 'gradient')).toStrictEqual({
			fillMode: 'gradient',
			fillGradientType: 'linear',
			fillGradientAngle: 90,
			fillGradientStops: [
				{ color: '#ff0000', position: 0, opacity: 1 },
				{ color: '#ffffff', position: 100, opacity: 1 },
			],
		});
		expect(shapeFillModeChange(style, 'pattern')).toStrictEqual({
			fillMode: 'pattern',
			fillPatternPreset: 'pct20',
			fillPatternBackgroundColor: '#ffffff',
		});
		expect(shapeFillModeChange(style, 'image')).toStrictEqual({
			fillMode: 'image',
			fillImageMode: 'stretch',
		});
		expect(shapeFillModeChange(style, 'background')).toStrictEqual({
			useBackgroundFill: true,
			fillMode: 'none',
			fillColor: 'transparent',
		});
		expect(shapeFillModeChange({ useBackgroundFill: true }, 'solid')).toMatchObject({
			useBackgroundFill: false,
			fillMode: 'solid',
		});
	});

	it('keeps the theme colours of existing stops when switching to a gradient', () => {
		const ref = { kind: 'scheme', name: 'accent1' } as unknown as NonNullable<
			ShapeStyle['fillColorRef']
		>;
		const stops = [
			{ color: '#ffffff', position: 100 },
			{ color: '#000000', position: 0, colorRef: ref },
		];
		expect(
			shapeFillModeChange({ fillGradientStops: stops }, 'gradient').fillGradientStops,
		).toStrictEqual([stops[1], stops[0]]);
	});

	it('writes transparency, gradient type, angle and tiling as style patches', () => {
		const style: ShapeStyle = { fillMode: 'gradient', fillOpacity: 0.5 };
		const patch = (field: string, raw: string) =>
			inspectorFieldPatch(SHAPE_FILL_OPTIONS_SECTIONS, 'fillOptions', field, style, raw);
		expect(resolveInspectorSections(SHAPE_FILL_OPTIONS_SECTIONS, style)[0]!.fields[0]!.value).toBe(
			50,
		);
		expect(patch('fillOpacity', '25')).toStrictEqual({ fillOpacity: 0.25 });
		expect(patch('fillGradientType', 'radial')).toStrictEqual({
			fillMode: 'gradient',
			fillGradientType: 'radial',
		});
		expect(patch('fillGradientAngle', '370')).toStrictEqual({
			fillMode: 'gradient',
			fillGradientAngle: 10,
		});
		expect(patch('fillImageMode', 'tile')).toStrictEqual({ fillImageMode: 'tile' });
	});
});
