import type { ShapeStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { resolveShapeFillView, shapeFillPatch } from './shape-fill-actions';

const ref = { kind: 'scheme', name: 'accent2' } as unknown as NonNullable<
	ShapeStyle['fillColorRef']
>;

const gradient: ShapeStyle = {
	fillMode: 'gradient',
	fillGradientStops: [
		{ color: '#000000', position: 0, opacity: 0.5, colorRef: ref },
		{ color: '#ffffff', position: 100 },
	],
};

describe('shape fill view', () => {
	it('shows the solid colour picker with its theme colour', () => {
		expect(resolveShapeFillView({ fillColor: '#112233', fillColorRef: ref })).toStrictEqual({
			mode: 'solid',
			color: { value: '#112233', colorRef: ref },
		});
		expect(resolveShapeFillView({ fillMode: 'none' })).toStrictEqual({ mode: 'none' });
		expect(resolveShapeFillView({ useBackgroundFill: true })).toStrictEqual({
			mode: 'background',
		});
	});

	it('lists the gradient stops in order, the last two not removable', () => {
		expect(resolveShapeFillView(gradient).stops).toStrictEqual([
			{ index: 0, color: '#000000', colorRef: ref, position: 0, opacity: 50, removable: false },
			{ index: 1, color: '#ffffff', position: 100, opacity: 100, removable: false },
		]);
	});

	it('draws the pattern swatches in the shape colours', () => {
		const view = resolveShapeFillView({
			fillMode: 'pattern',
			fillColor: '#ff0000',
			fillPatternPreset: 'cross',
		}).pattern!;
		expect(view.foreground).toBe('#ff0000');
		expect(view.background).toBe('#ffffff');
		expect(view.swatches).toHaveLength(56);
		expect(view.swatches.filter((s) => s.selected).map((s) => s.value)).toStrictEqual(['cross']);
		expect(view.swatches.every((s) => s.labelKey in translationsEn)).toBe(true);
		expect(view.swatches[0]!.backgroundImage).toMatch(/^url\('data:image\/svg\+xml;utf8,/);
	});

	it('shows the picture source', () => {
		expect(resolveShapeFillView({ fillMode: 'image', fillImageUrl: 'data:x' }).image).toStrictEqual(
			{ url: 'data:x' },
		);
	});
});

describe('shape fill actions', () => {
	it('commits colours as the ribbon does, with or without a theme colour', () => {
		expect(
			shapeFillPatch({ fillMode: 'pattern' }, { type: 'color', hex: '#abcdef' }),
		).toStrictEqual({
			fillColor: '#abcdef',
			fillColorRef: undefined,
			fillMode: 'solid',
		});
		expect(shapeFillPatch({}, { type: 'patternForeground', hex: '#010203', ref })).toStrictEqual({
			fillMode: 'pattern',
			fillColor: '#010203',
			fillColorRef: ref,
		});
		expect(shapeFillPatch({}, { type: 'patternBackground', hex: '#eeeeee' })).toStrictEqual({
			fillMode: 'pattern',
			fillPatternBackgroundColor: '#eeeeee',
		});
		expect(shapeFillPatch({}, { type: 'patternPreset', preset: 'dkHorz' })).toStrictEqual({
			fillMode: 'pattern',
			fillPatternPreset: 'dkHorz',
		});
	});

	it('leaves a slide background fill when the user picks a colour', () => {
		expect(
			shapeFillPatch({ useBackgroundFill: true }, { type: 'color', hex: '#000000' }),
		).toMatchObject({ useBackgroundFill: false, fillMode: 'solid' });
	});

	it('edits stops and keeps the theme colours of the others', () => {
		const clamped = shapeFillPatch(gradient, { type: 'stopPosition', index: 1, position: 140 });
		expect(clamped!.fillGradientStops![1]!.position).toBe(100);
		const moved = shapeFillPatch(
			{
				...gradient,
				fillGradientStops: [...gradient.fillGradientStops!, { color: '#00ff00', position: 60 }],
			},
			{ type: 'stopPosition', index: 0, position: 80 },
		);
		expect(moved!.fillGradientStops!.map((s) => [s.color, s.position, s.colorRef])).toStrictEqual([
			['#00ff00', 60, undefined],
			['#000000', 80, ref],
			['#ffffff', 100, undefined],
		]);
		const recoloured = shapeFillPatch(gradient, { type: 'stopColor', index: 0, hex: '#ff0000' });
		expect(recoloured!.fillGradientStops![0]).toStrictEqual({
			color: '#ff0000',
			position: 0,
			opacity: 0.5,
			colorRef: undefined,
		});
		const faded = shapeFillPatch(gradient, { type: 'stopOpacity', index: 1, opacity: 140 });
		expect(faded!.fillGradientStops![1]!.opacity).toBe(1);
		expect(shapeFillPatch(gradient, { type: 'stopOpacity', index: 1, opacity: NaN })).toBe(
			undefined,
		);
		expect(shapeFillPatch(gradient, { type: 'stopColor', index: 7, hex: '#000000' })).toBe(
			undefined,
		);
	});

	it('adds a stop halfway along the last gap and keeps two at least', () => {
		const added = shapeFillPatch(gradient, { type: 'addStop' })!;
		expect(added.fillMode).toBe('gradient');
		expect(added.fillGradientStops!.map((s) => [s.color, s.position])).toStrictEqual([
			['#000000', 0],
			['#000000', 50],
			['#ffffff', 100],
		]);
		expect(shapeFillPatch(gradient, { type: 'removeStop', index: 0 })).toBe(undefined);
		const three = { ...gradient, fillGradientStops: added.fillGradientStops };
		expect(
			shapeFillPatch(three, { type: 'removeStop', index: 1 })!.fillGradientStops!.map(
				(s) => s.position,
			),
		).toStrictEqual([0, 100]);
	});

	it('sets the picture source and mode switches through the fill kinds', () => {
		expect(shapeFillPatch({}, { type: 'imageUrl', url: 'data:y' })).toStrictEqual({
			fillMode: 'image',
			fillImageUrl: 'data:y',
		});
		expect(shapeFillPatch({}, { type: 'mode', mode: 'none' })).toStrictEqual({
			fillMode: 'none',
			fillColor: 'transparent',
		});
	});
});
