// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { applyAnimationOpacity, getAnimationOpacity } from './animation-opacity-style';

it('applies fixed opacity and restores authored alpha independently from CSS ramps', () => {
	const element = document.createElement('div');
	element.style.opacity = '.6';
	element.style.animation = 'ramp 1s both';
	const state = { visible: true, cssAnimation: undefined, opacity: 0.4 };
	expect(getAnimationOpacity(state, 0.6)).toBe(0.4);
	applyAnimationOpacity(element, state);
	expect(element.style.opacity).toBe('0.4');
	expect(element.style.animation).toBe('ramp 1s both');
	applyAnimationOpacity(element, undefined);
	expect(element.style.opacity).toBe('0.6');
});
