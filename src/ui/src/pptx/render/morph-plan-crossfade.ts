/**
 * Move a grouped crossfade pair's dissolve off its two elements and onto the
 * wrappers the binding paints each half in (see {@link liftCrossfadeToWrapper}).
 *
 * @module render/morph-plan-crossfade
 */
import { splitCssList } from './css-top-level-split';
import type { MorphCrossfadeGroup } from './morph-crossfade-group';

/**
 * Move a grouped pair's dissolve off the two elements and onto their wrappers.
 *
 * A pair that dissolves where it stands does not move, so the journey keyframes
 * it was given are `translate(0, 0) scale(1, 1) rotate(0deg)` at both ends. They
 * animate nothing and cost a compositing layer, whose raster the browser snaps
 * to whole device pixels - the wording is then painted a fraction of a pixel
 * from where the live stage paints it and twitches as the overlay comes and
 * goes. Dropping them and running only the fade, on the slide-sized wrapper the
 * binding puts each half in, measured pixel-identical to the settled slide
 * (issue #161; the same reasoning already governs an inert pair's ghost).
 *
 * A pair that MOVES keeps everything on the element: it needs the journey, and
 * a travelling shape cannot be seen to snap.
 */
export function liftCrossfadeToWrapper(
	group: MorphCrossfadeGroup,
	animations: readonly {
		elementId: string;
		animation: string;
		keyframes: string;
		target?: string;
	}[],
	outgoingAnimations: Map<string, string>,
	overlayIncomingAnimations: Map<string, string>,
): MorphCrossfadeGroup {
	// One entry can carry several tracks in its shorthand (a journey and its
	// fade), so they are separated and classified against their own keyframes.
	// A track that only touches opacity is the dissolve. One that touches
	// transform is the journey, and it only comes off the element when it says
	// the SAME thing at both ends - which is what an in-slot dissolve emits, and
	// what an inert pair's ghost already has dropped for the same reason. The
	// pair's boxes are no help here: a paragraph re-fitted around new wording
	// sits at a different height on the two slides and still does not move,
	// because `morphTextReplacedInSlot` refuses to interpolate it. Anything else
	// (an `<img>` crop) means the pair is doing more than dissolving.
	const tracksOf = (elementId: string): { fade: string[]; travels: boolean; other: number } => {
		const fade: string[] = [];
		let travels = false;
		let other = 0;
		for (const animation of animations) {
			if (animation.elementId !== elementId) {
				continue;
			}
			if (animation.target !== undefined) {
				other += 1;
				continue;
			}
			for (const track of splitCssList(animation.animation)) {
				const block = keyframesBlockOf(track.split(/\s+/u)[0], animation.keyframes);
				if (block === undefined) {
					other += 1;
				} else if (block.includes('transform')) {
					travels ||= !isStationaryKeyframes(block);
				} else if (block.includes('opacity')) {
					fade.push(track);
				} else {
					other += 1;
				}
			}
		}
		return { fade, travels, other };
	};
	const outgoing = tracksOf(group.outgoing.id);
	const incoming = tracksOf(group.incoming.id);
	if (
		outgoing.fade.length === 0 ||
		incoming.fade.length === 0 ||
		outgoing.travels ||
		incoming.travels ||
		outgoing.other + incoming.other > 0
	) {
		return group;
	}
	// Neutralised rather than dropped. A binding whose live stage is a SIBLING of
	// the overlay (Angular) has to emit its `incoming` rules unscoped, and the
	// one holding the stage copy invisible then matches the overlay's copy of the
	// same element too. Without a scoped rule to outrank it the arriving half
	// never painted at all, which is issue #160's defect exactly. `none` is that
	// rule, and it leaves the dissolve to the wrapper.
	outgoingAnimations.set(group.outgoing.id, 'none');
	overlayIncomingAnimations.set(group.incoming.id, 'none');
	return {
		...group,
		outgoingAnimation: outgoing.fade.join(', '),
		incomingAnimation: incoming.fade.join(', '),
	};
}

/**
 * Whether a keyframes block moves nothing: every frame states the same
 * transform, so running it is pure cost - a compositing layer whose raster the
 * browser snaps to whole device pixels (issue #161).
 */
function isStationaryKeyframes(block: string): boolean {
	// Compared on VALUE, not text: the generator writes the same standstill as
	// `translate(0px, 0px)` at one end and `translate(0, 0)` at the other, and a
	// string compare would call that a journey.
	const normalise = (value: string): string =>
		value
			.replace(/\s+/gu, '')
			.replace(/(-?\d*\.?\d+)(px|deg|%)?/gu, (_match, number: string, unit = '') =>
				Number(number) === 0 ? '0' : `${Number(number)}${unit}`,
			);
	const values = new Set(
		[...block.matchAll(/(?<![\w-])transform:\s*([^;}]+)/gu)].map((match) => normalise(match[1])),
	);
	return values.size <= 1;
}

/** The body of one `@keyframes` block inside a generated stylesheet. */
function keyframesBlockOf(name: string, keyframes: string): string | undefined {
	const start = keyframes.indexOf(`@keyframes ${name}`);
	if (start < 0) {
		return undefined;
	}
	let depth = 0;
	for (let index = keyframes.indexOf('{', start); index < keyframes.length; index += 1) {
		if (keyframes[index] === '{') {
			depth += 1;
		} else if (keyframes[index] === '}') {
			depth -= 1;
			if (depth === 0) {
				return keyframes.slice(start, index + 1);
			}
		}
	}
	return undefined;
}
