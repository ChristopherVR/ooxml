import type { PptxElement } from 'ooxml-core/pptx';
import { expect, it } from 'vitest';
import {
	getBackgroundAnimationLayers,
	getBackgroundAnimationState,
} from './background-animation-layers';

const shape = { id: 's', type: 'shape' } as PptxElement;
it('seeds background fill and stroke colour variables on paint, keeping whole opacity on the box', () => {
	const element = {
		...shape,
		shapeStyle: { fillColor: '#00aa00', strokeColor: '#0000ff' },
	} as PptxElement;
	const layers = getBackgroundAnimationLayers(
		element,
		{ opacity: 0.4, animation: 'whole-opacity 1s both' },
		{
			visible: true,
			cssAnimation: 'paint-colour 1s both',
			animatesFill: true,
			animatesStroke: true,
		},
	);
	expect(layers.box).toEqual({ opacity: 0.4, animation: 'whole-opacity 1s both' });
	expect(layers.paint).toMatchObject({
		'--pptx-animation-fill-base': '#00aa00',
		'--pptx-animation-stroke-base': '#0000ff',
		animation: 'paint-colour 1s both',
	});
	expect(layers.paint).not.toHaveProperty('opacity');
});
it('separates animated paint from positioning without hiding or moving text', () => {
	const style = {
		left: '10px',
		width: '200px',
		transform: 'rotate(10deg)',
		backgroundColor: 'red',
		borderColor: 'black',
		opacity: 0.6,
	};
	const state = { visible: false, cssAnimation: 'fade 1s both' };
	const layers = getBackgroundAnimationLayers(shape, style, state);
	expect(layers.box).toEqual({
		left: '10px',
		width: '200px',
		transform: 'rotate(10deg)',
		opacity: 0.6,
	});
	expect(layers.paint).toMatchObject({
		backgroundColor: 'red',
		borderColor: 'black',
		visibility: 'hidden',
		animation: 'fade 1s both',
	});
	expect(layers.targetId).toBe('s::pptx-bg');
	expect(layers.paint).not.toHaveProperty('opacity');
	expect(style.backgroundColor).toBe('red');
	expect(getBackgroundAnimationLayers(shape, style, undefined).box).toBe(style);
});
it('uses only supported shape and text background targets', () => {
	const state = { visible: false, cssAnimation: undefined };
	const states = new Map([['s::pptx-bg', state]]);
	expect(getBackgroundAnimationState(shape, states)).toBe(state);
	expect(getBackgroundAnimationState({ ...shape, type: 'text' } as PptxElement, states)).toBe(
		state,
	);
	expect(
		getBackgroundAnimationState({ ...shape, type: 'chart' } as PptxElement, states),
	).toBeUndefined();
});
