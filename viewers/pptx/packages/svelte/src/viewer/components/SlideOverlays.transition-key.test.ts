import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

/** The component's own text: mounting it needs the whole editor stack. */
const source = readFileSync(
	[
		'src/viewer/components/SlideOverlays.svelte',
		'packages/svelte/src/viewer/components/SlideOverlays.svelte',
	]
		.map((candidate) => resolve(process.cwd(), candidate))
		.find((candidate) => existsSync(candidate))!,
	'utf8',
);

// Issue #141, slides 6 -> 7: slide 6 advances after 0 s, so its timer fires as
// its fade ends, while that overlay is still mounted. The overlay arms its
// completion timer on mount, so a reused instance let the fade's timer tear
// down slide 7's morph 50 ms in and slide 7 appeared in one frame.
describe('slide transition overlay mounting', () => {
	it('remounts the overlay for every transition', () => {
		const block = source.slice(source.indexOf('{#if presenting && presentationTransition}'));
		const key = block.indexOf('{#key presentationTransition}');
		const overlay = block.indexOf('<PresentationTransitionOverlay');
		expect(key).toBeGreaterThan(0);
		expect(key).toBeLessThan(overlay);
		expect(block.indexOf('{/key}')).toBeGreaterThan(overlay);
	});
});
