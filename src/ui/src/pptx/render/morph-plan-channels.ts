/**
 * Sort the flat list `morph-animation` generates into the channels of a plan:
 * container animations by half, the `<img>` crop channel, the outline,
 * custom-geometry and connector-strand channels (see {@link MorphTransitionPlan}).
 *
 * @module render/morph-plan-channels
 */
import type { generateFullMorphTransition } from './morph-animation';

/** The animations `morph-animation` generated for one transition. */
export type MorphGeneratedAnimations = ReturnType<typeof generateFullMorphTransition>;

/** The generated animations, partitioned by the channel a binding applies them on. */
export interface MorphAnimationChannels {
	/** Every `@keyframes` block, in generation order. */
	keyframes: string[];
	incomingAnimations: Map<string, string>;
	outgoingAnimations: Map<string, string>;
	incomingImageAnimations: Map<string, string>;
	outgoingImageAnimations: Map<string, string>;
	outlineAnimations: Map<string, string>;
	outlineStrandAnimations: Map<string, string[]>;
	customGeometryAnimations: Map<string, string>;
	lineAnimations: Map<string, string>;
}

/**
 * Partition `animations` into channels. `outgoingIds` is the id set of the
 * outgoing slide's (flattened) elements: partitioning on it, rather than on id
 * shape, keeps the id format an implementation detail of core.
 */
export function partitionMorphAnimations(
	animations: MorphGeneratedAnimations,
	outgoingIds: ReadonlySet<string>,
): MorphAnimationChannels {
	const incomingAnimations = new Map<string, string>();
	const outgoingAnimations = new Map<string, string>();
	const incomingImageAnimations = new Map<string, string>();
	const outgoingImageAnimations = new Map<string, string>();
	const outlineAnimations = new Map<string, string>();
	const outlineStrandAnimations = new Map<string, string[]>();
	const customGeometryAnimations = new Map<string, string>();
	const lineAnimations = new Map<string, string>();
	const keyframes: string[] = [];

	for (const animation of animations) {
		keyframes.push(animation.keyframes);
		const isOutgoing = outgoingIds.has(animation.elementId);
		// An `image`-targeted animation rides the `<img>` inside the element, so
		// it goes in its own map: it shares an element id with the container
		// animation and would otherwise overwrite it.
		if (animation.target === 'outline' && animation.strand !== undefined) {
			const strands = outlineStrandAnimations.get(animation.elementId) ?? [];
			strands[animation.strand] = animation.animation;
			outlineStrandAnimations.set(animation.elementId, strands);
			continue;
		}
		if (animation.target === 'outline' || animation.target === 'custom') {
			(animation.target === 'outline' ? outlineAnimations : customGeometryAnimations).set(
				animation.elementId,
				animation.animation,
			);
			continue;
		}
		if (animation.target === 'line') {
			lineAnimations.set(animation.elementId, animation.animation);
			continue;
		}
		const target =
			animation.target === 'image'
				? isOutgoing
					? outgoingImageAnimations
					: incomingImageAnimations
				: isOutgoing
					? outgoingAnimations
					: incomingAnimations;
		// One element can be driven by SEVERAL animations at once, so they are
		// composed into the shorthand's comma-separated list rather than
		// overwriting each other. A pair whose shape type or adjustment values
		// changed gets both its journey (`transform`, from
		// `generateMorphAnimations`) and its outline tween (`clip-path`, from
		// `generateGeometryMorphAnimation`), keyed on the same incoming id; a
		// plain `set` dropped whichever was generated first, so a rect gliding
		// into a distant ellipse stopped travelling altogether and sat at its
		// destination re-cutting its own outline while its ghost flew over.
		//
		// Composing is only correct because the animations here touch DISJOINT
		// properties. Anything that would fight over one property has to keep
		// using its own channel, the way the `a:srcRect` crop does above.
		const existing = target.get(animation.elementId);
		target.set(
			animation.elementId,
			existing === undefined ? animation.animation : `${existing}, ${animation.animation}`,
		);
	}
	return {
		keyframes,
		incomingAnimations,
		outgoingAnimations,
		incomingImageAnimations,
		outgoingImageAnimations,
		outlineAnimations,
		outlineStrandAnimations,
		customGeometryAnimations,
		lineAnimations,
	};
}
