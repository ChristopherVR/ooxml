// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { applyAnimationOpacity, getAnimationOpacity } from './animation-opacity-style';

it('applies fixed opacity and restores authored alpha independently from CSS ramps', () => {
	const element = document.createElement('div');
	element.style.opacity = '.6';
	element.style.animation = 'ramp 1s both';
	// The engine's own serialisation of the ramp (value order differs between engines).
	const authoredAnimation = element.style.animation;
	const state = { visible: true, cssAnimation: undefined, opacity: 0.4 };
	expect(getAnimationOpacity(state, 0.6)).toBe(0.4);
	applyAnimationOpacity(element, state);
	expect(element.style.opacity).toBe('0.4');
	expect(element.style.animation).toBe(authoredAnimation);
	expect(element.style.animationName).toBe('ramp');
	applyAnimationOpacity(element, undefined);
	expect(element.style.opacity).toBe('0.6');
});
