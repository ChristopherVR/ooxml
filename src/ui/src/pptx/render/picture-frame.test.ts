import type { PptxElement, ShapeStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	getPictureFrameLayers,
	getPictureShadowFilter,
	hasPictureFrame,
	getShapeEffectLayerPlacement,
	getShapeEffectAnimationLayerStyle,
} from './picture-frame';

/** The frame a Picture Styles pick writes: a centred red line and an outer shadow. */
const FRAMED: ShapeStyle = {
	strokeWidth: 10.67,
	strokeFillMode: 'solid',
	strokeColor: '#FF0000',
	shadowColor: '#000000',
	shadowOpacity: 0.6,
	shadowBlur: 5.33,
	shadowAngle: 45,
	shadowDistance: 10.67,
};

function picture(shapeStyle?: ShapeStyle): PptxElement {
	return {
		type: 'picture',
		id: 'pic-1',
		x: 0,
		y: 0,
		width: 200,
		height: 150,
		imageData: 'data:image/png;base64,AA==',
		...(shapeStyle ? { shapeStyle } : {}),
	} as PptxElement;
}

const CLIP = { overflow: 'hidden', borderRadius: '12px', clipPath: 'inset(0)' } as const;

it('keeps shadow paint on an animated layer without mutating the static style', () => {
	const original = { filter: 'drop-shadow(1px 2px 3px black)', clipPath: 'inset(0)' };
	const style = getShapeEffectAnimationLayerStyle(picture(FRAMED), original, {
		visible: false,
		cssAnimation: 'fade 1s both',
	});
	expect(style).toMatchObject({
		...original,
		clipPath: 'none',
		overflow: 'visible',
		position: 'absolute',
		inset: '0',
		visibility: 'hidden',
		animation: 'fade 1s both',
	});
	expect(original).toEqual({ filter: 'drop-shadow(1px 2px 3px black)', clipPath: 'inset(0)' });
});

it('places picture effects above pixels and shape effects below text on every surface', () => {
	expect(getShapeEffectLayerPlacement(picture(FRAMED))).toEqual({
		beforeBody: false,
		afterBody: true,
	});
	expect(getShapeEffectLayerPlacement({ ...picture(), type: 'shape' } as PptxElement)).toEqual({
		beforeBody: true,
		afterBody: false,
	});
});

describe('getPictureFrameLayers', () => {
	it('keeps the clip on the box when the picture has no frame', () => {
		expect(getPictureFrameLayers(picture(), CLIP)).toEqual({
			box: { overflow: 'hidden', borderRadius: '12px', clipPath: 'inset(0)' },
			surface: undefined,
		});
	});

	it('moves the clip to a surface and unclips the box for a framed picture', () => {
		const layers = getPictureFrameLayers(picture(FRAMED), CLIP);
		expect(layers.surface).toEqual({
			position: 'absolute',
			inset: '0',
			overflow: 'hidden',
			borderRadius: '12px',
			clipPath: 'inset(0)',
		});
		expect(layers.box.overflow).toBe('visible');
		expect(layers.box.filter).toMatch(/^drop-shadow\(/);
	});

	it('splits for an outline alone, with no shadow filter', () => {
		const { strokeWidth, strokeFillMode, strokeColor } = FRAMED;
		const layers = getPictureFrameLayers(
			picture({ strokeWidth, strokeFillMode, strokeColor }),
			CLIP,
		);
		expect(layers.surface).toBeDefined();
		expect(layers.box).toEqual({ overflow: 'visible' });
	});

	it('splits for a shadow alone', () => {
		const { strokeWidth: _w, strokeFillMode: _m, strokeColor: _c, ...shadowOnly } = FRAMED;
		expect(getPictureFrameLayers(picture(shadowOnly), CLIP).surface).toBeDefined();
	});
});

describe('hasPictureFrame', () => {
	it('ignores a width-only line with no paint (a frameless picture)', () => {
		expect(hasPictureFrame(picture({ strokeWidth: 1 }))).toBe(false);
	});

	it('is false for a shape, which paints its own outline elsewhere', () => {
		expect(hasPictureFrame({ ...picture(FRAMED), type: 'shape' } as PptxElement)).toBe(false);
		expect(getPictureShadowFilter({ ...picture(FRAMED), type: 'shape' } as PptxElement)).toBe(
			undefined,
		);
	});
});
