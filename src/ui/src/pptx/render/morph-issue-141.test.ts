/**
 * Issue #141's morph facts, on the slides of the reporter's deck rebuilt as
 * the plan sees them (`__fixtures__/morph-issue-141`): what pairs, what fades,
 * how the overlay stacks a dissolving column between the picture bands, which
 * way the 9 -> 10 heading travels, and that the parked paragraph of 17 -> 18
 * is painted once in either direction. The narrower tests next to each fix
 * isolate one rule; this one holds them together on the real arrangement.
 */
import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { matchMorphElementsFull } from './morph-matching';
import {
	SLIDE_10,
	SLIDE_17,
	SLIDE_18,
	SLIDE_6,
	SLIDE_7,
	SLIDE_9,
} from './__fixtures__/morph-issue-141';
import { morphOverlayElementSlots } from './morph-overlay-stack';
import type { MorphTransitionPlan } from './morph-plan';
import { buildMorphTransitionPlan } from './morph-plan';

const DURATION = 2000;

const plan = (from: PptxSlide, to: PptxSlide): MorphTransitionPlan =>
	buildMorphTransitionPlan(from, to, DURATION, 'object')!;

const pairsOf = (from: PptxSlide, to: PptxSlide): string[] =>
	matchMorphElementsFull(from, to).pairs.map(
		(pair) => `${pair.fromElement.id}>${pair.toElement.id}`,
	);

/** The body of `@keyframes name`, or undefined. */
function keyframes(css: string, name: string): string | undefined {
	return css.split(`@keyframes ${name} {`)[1]?.split('@keyframes')[0];
}

/** Every opacity stop of a keyframes body, as [percent, opacity]. */
function opacityStops(body: string): [number, number][] {
	const stops: [number, number][] = [];
	for (const match of body.matchAll(/([\d.%,\sfromto]+)\{([^}]*)\}/gu)) {
		const opacity = /opacity:\s*([\d.]+)/u.exec(match[2]!)?.[1];
		if (opacity === undefined) {
			continue;
		}
		for (const label of match[1]!.split(',')) {
			const key = label.trim();
			const percent = key === 'from' ? 0 : key === 'to' ? 100 : Number.parseFloat(key);
			if (Number.isFinite(percent)) {
				stops.push([percent, Number(opacity)]);
			}
		}
	}
	return stops.sort((a, b) => a[0] - b[0]);
}

/**
 * The opacity an `animation` shorthand list holds a copy at, at `percent` of
 * the morph: each named track's stops interpolated linearly (the timing
 * curve never turns a 0 -> 0 hold visible), multiplied together.
 */
function opacityAt(css: string, animation: string | undefined, percent: number): number {
	let opacity = 1;
	for (const track of (animation ?? '').split(',')) {
		const name = track.trim().split(/\s+/u)[0];
		const body = name ? keyframes(css, name) : undefined;
		const stops = body ? opacityStops(body) : [];
		if (stops.length === 0) {
			continue;
		}
		const after = stops.findIndex(([at]) => at >= percent);
		if (after <= 0) {
			opacity *= stops[after === 0 ? 0 : stops.length - 1]![1];
			continue;
		}
		const [a, va] = stops[after - 1]!;
		const [b, vb] = stops[after]!;
		opacity *= b === a ? vb : va + ((vb - va) * (percent - a)) / (b - a);
	}
	return opacity;
}

/** The copies of a pair painted at `percent`: its overlay ghost and its stage copy. */
function visibleCopies(
	result: MorphTransitionPlan,
	fromId: string,
	toId: string,
	percent: number,
): string[] {
	const visible: string[] = [];
	if (result.outgoingElements.some((element) => element.id === fromId)) {
		if (opacityAt(result.keyframesCss, result.outgoingAnimations.get(fromId), percent) > 0.01) {
			visible.push(`ghost ${fromId}`);
		}
	}
	if (opacityAt(result.keyframesCss, result.incomingAnimations.get(toId), percent) > 0.01) {
		visible.push(`stage ${toId}`);
	}
	return visible;
}

describe('issue #141, slides 6 -> 7: dissolving "L" columns under the picture bands', () => {
	const result = plan(SLIDE_6, SLIDE_7);

	it('carries the three columns as dissolving pairs', () => {
		expect(
			result.crossfadeGroups.map((group) => `${group.outgoing.id}>${group.incoming.id}`),
		).toStrictEqual(['s6-shape-3>s7-shape-4', 's6-shape-4>s7-shape-5', 's6-shape-5>s7-shape-6']);
	});

	it('pairs the backdrop, the strip and the top band by their pictures', () => {
		const pairs = pairsOf(SLIDE_6, SLIDE_7);
		expect(pairs).toEqual(
			expect.arrayContaining(['s6-pic-0>s7-pic-0', 's6-pic-1>s7-pic-1', 's6-pic-3>s7-pic-3']),
		);
		// The bottom band changes picture, so it fades out and the new one in.
		expect(pairs.some((pair) => pair.startsWith('s6-pic-2>'))).toBe(false);
	});

	it('stacks the top band and the arriving bottom band above the columns, the strip below', () => {
		const slots = morphOverlayElementSlots(result);
		const lowestGroup = Math.min(...slots.groups);
		const highestGroup = Math.max(...slots.groups);
		expect(slots.groups).toHaveLength(3);
		expect(slots.elements.get('s6-pic-1')).toBeLessThan(lowestGroup);
		expect(slots.elements.get('s6-pic-3')).toBeGreaterThan(highestGroup);
		expect(slots.elements.get('s7-pic-2')).toBeGreaterThan(highestGroup);
	});

	it('fades in only the upper of the two stacked "L" copies', () => {
		const lower = result.incomingAnimations.get('s7-shape-3')?.split(' ')[0];
		const upper = result.incomingAnimations.get('s7-shape-7')?.split(' ')[0];
		expect(keyframes(result.keyframesCss, lower!)).toMatch(/0%, 99\.9% \{\s*opacity: 0;/u);
		expect(keyframes(result.keyframesCss, upper!)).not.toContain('99.9%');
	});
});

describe('issue #141, slides 9 -> 10: the heading drops from the top', () => {
	it('pairs the heading by its words, not the label sharing its new shape id', () => {
		const pairs = pairsOf(SLIDE_9, SLIDE_10);
		// Shape 24 lands as shape 27, which the outgoing slide's label also uses.
		expect(pairs).toContain('s9-shape-1>s10-shape-3');
		expect(pairs).toContain('s9-shape-4>s10-shape-6');
	});

	it('travels straight down', () => {
		const result = plan(SLIDE_9, SLIDE_10);
		const name = result.incomingAnimations.get('s10-shape-3')?.split(' ')[0];
		const [, dx, dy] =
			/translate\((-?[\d.]+)px, (-?[\d.]+)px\)/u.exec(keyframes(result.keyframesCss, name!)!) ?? [];
		expect(Number(dx)).toBe(0);
		expect(Number(dy)).toBeLessThan(-300);
	});
});

describe('issue #141, slides 17 <-> 18: columns that trade paragraphs and a parked one', () => {
	it('fades both columns in place and carries the parked paragraph', () => {
		const match = matchMorphElementsFull(SLIDE_17, SLIDE_18);
		const ids = (elements: { id: string }[]): string[] => elements.map((element) => element.id);
		expect(ids(match.unmatchedFrom)).toStrictEqual(['s17-shape-0', 's17-shape-2']);
		expect(ids(match.unmatchedTo)).toStrictEqual(['s18-shape-0', 's18-shape-1']);
		expect(pairsOf(SLIDE_17, SLIDE_18)).toContain('s17-shape-3>s18-shape-2');
	});

	it('paints the parked paragraph once forward, in its new layout', () => {
		const result = plan(SLIDE_17, SLIDE_18);
		expect(result.outgoingTextReplacements?.get('s17-shape-3')).toBe('s18-shape-2');
		for (const percent of [25, 50, 75]) {
			expect(visibleCopies(result, 's17-shape-3', 's18-shape-2', percent)).toStrictEqual([
				'ghost s17-shape-3',
			]);
		}
	});

	it('paints it once backward, in its own layout', () => {
		const result = plan(SLIDE_18, SLIDE_17);
		expect(result.outgoingTextReplacements?.has('s18-shape-2')).toBe(false);
		for (const percent of [25, 50, 75]) {
			expect(visibleCopies(result, 's18-shape-2', 's17-shape-3', percent)).toStrictEqual([
				'ghost s18-shape-2',
			]);
		}
	});
});
