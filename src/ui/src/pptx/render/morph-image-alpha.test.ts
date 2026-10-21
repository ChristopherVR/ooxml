import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { isInertMorphPair } from './morph-animation';
import { morphImageAlphaChanged } from './morph-image-alpha';
import { buildMorphTransitionPlan } from './morph-plan';

/**
 * A full-slide picture, the way issue #141's deck fades one in: the SAME blip
 * on both slides, `<a:alphaModFix amt="0"/>` on the first and a bare
 * `<a:alphaModFix/>` (fully opaque) on the second.
 */
function picture(id: string, imagePath: string, alphaModFix?: number): PptxElement {
	return {
		id,
		name: id,
		type: 'picture',
		x: 0,
		y: 0,
		width: 1280,
		height: 720,
		imagePath,
		...(alphaModFix === undefined ? {} : { imageEffects: { alphaModFix } }),
	} as PptxElement;
}

/** A band with no counterpart, which dissolves out beneath the faded picture. */
function band(id: string): PptxElement {
	return { ...picture(id, 'ppt/media/band.png'), y: 456, height: 264 } as PptxElement;
}

function slide(id: string, elements: PptxElement[]): PptxSlide {
	return { id, slideNumber: 1, elements } as unknown as PptxSlide;
}

/** The `@keyframes` block an animation shorthand names. */
function keyframesOf(css: string, animation: string): string {
	const name = animation.split(/\s+/u)[0];
	const start = css.indexOf(`@keyframes ${name} `);
	expect(start, `keyframes for ${name}`).toBeGreaterThanOrEqual(0);
	return css.slice(start, css.indexOf('\n}', start));
}

const BACKDROP = 'ppt/media/image53.jpeg';
const FADED = 'ppt/media/image62.jpeg';

describe('morphImageAlphaChanged', () => {
	it('sees a transparency change on an otherwise identical picture', () => {
		expect(morphImageAlphaChanged(picture('a', FADED, 0), picture('b', FADED, 100))).toBe(true);
	});

	it('treats a missing alphaModFix as fully opaque', () => {
		expect(morphImageAlphaChanged(picture('a', FADED), picture('b', FADED, 100))).toBe(false);
		expect(morphImageAlphaChanged(picture('a', FADED), picture('b', FADED, 22))).toBe(true);
	});
});

describe('issue #141: a picture faded in through its transparency', () => {
	it('is not an inert pair', () => {
		// Before the fix this pair compared equal on every axis the engine read,
		// so its ghost was painted statically at amt="0" and the incoming half was
		// held invisible: the picture appeared in one frame when the overlay came
		// down.
		expect(isInertMorphPair(picture('a', FADED, 0), picture('b', FADED, 100))).toBe(false);
	});

	it('tweens the incoming picture from the outgoing transparency', () => {
		const from = slide('s8', [picture('s8-back', BACKDROP), picture('s8-faded', FADED, 0)]);
		const to = slide('s9', [picture('s9-back', BACKDROP), picture('s9-faded', FADED, 100)]);
		const plan = buildMorphTransitionPlan(from, to, 2000);

		const imageTrack = plan?.incomingImageAnimations.get('s9-faded');
		expect(imageTrack).toBeDefined();
		const frames = keyframesOf(plan!.keyframesCss, imageTrack!);
		expect(frames).toMatch(/from \{\s*opacity: 0;/u);
		expect(frames).toMatch(/to \{\s*opacity: 1;/u);

		// Nothing holds the live picture's container invisible for the morph.
		const container = plan!.incomingAnimations.get('s9-faded');
		if (container !== undefined) {
			expect(keyframesOf(plan!.keyframesCss, container)).not.toMatch(/to \{[^}]*opacity: 0;/u);
		}
	});

	it('takes a painted ghost to the same final transparency', () => {
		// The band dissolving out underneath is what makes the overlay paint the
		// faded picture's ghost (see `resolveMorphGhostIds`), as on the deck's
		// slide 8 -> 9.
		const from = slide('s8', [
			picture('s8-back', BACKDROP),
			band('s8-band'),
			picture('s8-faded', FADED, 0),
		]);
		const to = slide('s9', [picture('s9-back', BACKDROP), picture('s9-faded', FADED, 100)]);
		const plan = buildMorphTransitionPlan(from, to, 2000);

		expect(plan?.outgoingElements.map((element) => element.id)).toContain('s8-faded');
		const ghostTrack = plan?.outgoingImageAnimations.get('s8-faded');
		expect(ghostTrack).toBeDefined();
		const frames = keyframesOf(plan!.keyframesCss, ghostTrack!);
		expect(frames).toMatch(/from \{\s*opacity: 0;/u);
		expect(frames).toMatch(/to \{\s*opacity: 1;/u);
		expect(plan?.incomingImageAnimations.get('s9-faded')).toBeDefined();
	});

	it('leaves a picture whose transparency did not change alone', () => {
		const from = slide('s8', [picture('s8-faded', FADED, 40)]);
		const to = slide('s9', [picture('s9-faded', FADED, 40)]);
		const plan = buildMorphTransitionPlan(from, to, 2000);
		expect(plan?.incomingImageAnimations.size ?? 0).toBe(0);
		expect(plan?.outgoingImageAnimations.size ?? 0).toBe(0);
	});
});
