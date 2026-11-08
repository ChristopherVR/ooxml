/**
 * The SDK's picture tile and stretch inputs: the model fields and documented
 * defaults they map to, and the values they refuse. Saving is covered in
 * `core/runtime/save-picture-fill.test.ts`.
 */
import { describe, expect, it } from 'vitest';

import { createImageElement, createShapeElement } from './ElementFactory';
import type { FillInput, ImageOptions } from './types';

const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
type ImageFill = Extract<FillInput, { type: 'image' }>;
const fill = (extra: Partial<ImageFill>): FillInput => ({ type: 'image', url: PNG, ...extra });

describe('picture fill input', () => {
	it('maps a plain fill to a stretch with no layout', () => {
		const style = createShapeElement('rect', { fill: fill({}) }).shapeStyle;
		expect(style).toMatchObject({ fillMode: 'image', fillImageUrl: PNG, fillImageMode: 'stretch' });
		expect(style?.fillImageLayout).toBeUndefined();
	});

	it('fills in the tile defaults', () => {
		const style = createShapeElement('rect', { fill: fill({ mode: 'tile' }) }).shapeStyle;
		expect(style?.fillImageMode).toBe('tile');
		expect(style?.fillImageLayout).toEqual({
			tileOffsetX: 0,
			tileOffsetY: 0,
			tileScaleX: 1,
			tileScaleY: 1,
			tileFlip: 'none',
			tileAlignment: 'tl',
		});
	});

	it('takes a tile without a mode as a tile', () => {
		const style = createShapeElement('rect', { fill: fill({ tile: { scaleX: 2 } }) }).shapeStyle;
		expect(style?.fillImageMode).toBe('tile');
		expect(style?.fillImageLayout?.tileScaleX).toBe(2);
	});

	it('maps stretch insets', () => {
		const style = createShapeElement('rect', {
			fill: fill({ mode: 'stretch', fillRect: { l: 0, t: 0.1, r: -0.2, b: 0.3 } }),
		}).shapeStyle;
		expect(style?.fillImageLayout).toEqual({
			fillRectLeft: 0,
			fillRectTop: 0.1,
			fillRectRight: -0.2,
			fillRectBottom: 0.3,
		});
	});

	it.each<[string, Partial<ImageFill>]>([
		['a tile and a fill rectangle together', { tile: {}, fillRect: { l: 0, t: 0, r: 0, b: 0 } }],
		['a tile with mode stretch', { mode: 'stretch', tile: {} }],
		['a fill rectangle with mode tile', { mode: 'tile', fillRect: { l: 0, t: 0, r: 0, b: 0 } }],
		['an unknown mode', { mode: 'cover' as 'tile' }],
		['a zero scale', { tile: { scaleX: 0 } }],
		['a negative scale', { tile: { scaleY: -1 } }],
		['a NaN scale', { tile: { scaleX: Number.NaN } }],
		['an infinite offset', { tile: { offsetX: Number.POSITIVE_INFINITY } }],
		['an unknown flip', { tile: { flip: 'z' as 'x' } }],
		['an unknown alignment', { tile: { alignment: 'center' as 'ctr' } }],
		['an inset past ten frames', { fillRect: { l: 10.5, t: 0, r: 0, b: 0 } }],
		['a missing inset', { fillRect: { l: 0, t: 0, r: 0 } as never }],
		['insets that leave no width', { fillRect: { l: 0.5, t: 0, r: 0.5, b: 0 } }],
		['insets that leave no height', { fillRect: { l: 0, t: 0.75, r: 0, b: 0.5 } }],
		['a URL that is not a data URL', { url: 'https://example.com/a.png' }],
		['an archive path', { url: 'ppt/media/image1.png' }],
	])('refuses %s', (_name, extra) => {
		expect(() => createShapeElement('rect', { fill: fill(extra) })).toThrow(RangeError);
	});
});

describe('picture input', () => {
	it('leaves a picture without options alone', () => {
		const image = createImageElement(PNG, {});
		expect(image.tileScaleX).toBeUndefined();
		expect(image.fillRectLeft).toBeUndefined();
	});

	it('maps a tile onto the picture fields', () => {
		const image = createImageElement(PNG, { tile: { offsetY: 3, alignment: 'b' } });
		expect(image).toMatchObject({
			tileOffsetX: 0,
			tileOffsetY: 3,
			tileScaleX: 1,
			tileScaleY: 1,
			tileFlip: 'none',
			tileAlignment: 'b',
		});
	});

	it('maps stretch insets onto the picture fields', () => {
		const image = createImageElement(PNG, { fillRect: { l: 0.1, t: 0.2, r: 0.3, b: 0.4 } });
		expect(image).toMatchObject({
			fillRectLeft: 0.1,
			fillRectTop: 0.2,
			fillRectRight: 0.3,
			fillRectBottom: 0.4,
		});
	});

	it.each<[string, ImageOptions]>([
		['a tile and a fill rectangle together', { tile: {}, fillRect: { l: 0, t: 0, r: 0, b: 0 } }],
		['a zero scale', { tile: { scaleY: 0 } }],
		['an inset below ten frames', { fillRect: { l: 0, t: -11, r: 0, b: 0 } }],
	])('refuses %s', (_name, options) => {
		expect(() => createImageElement(PNG, options)).toThrow(RangeError);
	});
});
