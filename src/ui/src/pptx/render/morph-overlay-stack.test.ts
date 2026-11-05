import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import type { MorphCrossfadeGroup } from './morph-crossfade-group';
import { buildMorphTransitionPlan } from './morph-plan';
import {
	morphOverlayElementSlots,
	morphOverlayStackSlots,
	resolveMorphOverlayStack,
} from './morph-overlay-stack';

const slide = (id: string, elements: PptxElement[]): PptxSlide =>
	({ id, slideNumber: 1, elements }) as unknown as PptxSlide;

function picture(id: string, name: string, y: number, height: number): PptxElement {
	return {
		id,
		name,
		type: 'picture',
		x: 0,
		y,
		width: 1280,
		height,
		imagePath: `ppt/media/${name}.png`,
	} as PptxElement;
}

/** One of issue #141's "L" columns: the same runs, another font, another place. */
function column(id: string, y: number, fontFamily: string): PptxElement {
	return {
		id,
		name: 'Object 14',
		type: 'text',
		x: 284,
		y,
		width: 51,
		height: 1066,
		text: 'L\nL\nL',
		textStyle: { color: '#FFFFFF', fontFamily },
		textSegments: [{ text: 'L\nL\nL', style: { fontFamily, color: '#FFFFFF' } }],
	} as PptxElement;
}

describe('resolveMorphOverlayStack', () => {
	const a = { id: 'a', type: 'shape', x: 0, y: 0, width: 10, height: 10 } as PptxElement;
	const b = { id: 'b', type: 'shape', x: 0, y: 0, width: 10, height: 10 } as PptxElement;
	const c = { id: 'c', type: 'shape', x: 0, y: 0, width: 10, height: 10 } as PptxElement;
	const groupOf = (incoming: PptxElement): MorphCrossfadeGroup => ({
		outgoing: { ...incoming, id: `${incoming.id}-old` } as PptxElement,
		incoming,
	});

	it('splits a run where a group belongs inside it', () => {
		const { layers, stack } = resolveMorphOverlayStack(
			[a, c],
			[],
			[groupOf(b)],
			[a, { ...b, id: 'b-old' } as PptxElement, c],
			[a, b, c],
			[
				{ fromElement: a, toElement: a },
				{ fromElement: { ...b, id: 'b-old' } as PptxElement, toElement: b },
				{ fromElement: c, toElement: c },
			],
		);
		expect(layers.map((layer) => layer.elements.map((element) => element.id))).toEqual([
			['a'],
			['c'],
		]);
		expect(stack).toEqual([
			{ kind: 'layer', index: 0 },
			{ kind: 'crossfade', index: 0 },
			{ kind: 'layer', index: 1 },
		]);
	});

	it('falls back to every run under every group for a plan without a stack', () => {
		const plan = {
			overlayLayers: [
				{ side: 'outgoing' as const, elements: [a, b] },
				{ side: 'incoming' as const, elements: [c] },
			],
			crossfadeGroups: [{}],
		};
		expect(morphOverlayStackSlots(plan)).toEqual({ layers: [0, 1], groups: [2] });
		const flat = morphOverlayElementSlots(plan);
		expect([...flat.elements]).toEqual([
			['a', 0],
			['b', 1],
			['c', 2],
		]);
		expect(flat.groups).toEqual([3]);
	});
});

// Issue #141, slides 6 -> 7: the column changes font, so it dissolves, and it
// slides under the top band the incoming slide stacks ABOVE it. The left strip
// moves from above the column to below it, which lifts the column (and so
// groups both its halves) into the overlay.
describe('issue #141: a dissolving column under a picture band', () => {
	const from = slide('6', [
		picture('bg6', 'backdrop', 0, 720),
		column('col6', -916, 'Calibri Light'),
		picture('strip6', 'strip', -1, 720),
		picture('band6', 'band', -1, 196),
	]);
	const to = slide('7', [
		picture('bg7', 'backdrop', 0, 720),
		picture('strip7', 'strip', -3, 720),
		column('col7', -158, 'Calibri'),
		picture('band7', 'band', -2, 196),
	]);
	const plan = buildMorphTransitionPlan(from, to, 2000)!;

	it('paints the column as a crossfade group', () => {
		expect(plan.crossfadeGroups.map((group) => group.incoming.id)).toEqual(['col7']);
	});

	it('stacks the band above the column and the strip below it', () => {
		const slots = morphOverlayStackSlots(plan);
		const layerOf = (id: string): number =>
			plan.overlayLayers.findIndex((layer) => layer.elements.some((element) => element.id === id));
		const group = slots.groups[0]!;
		expect(slots.layers[layerOf('band6')]).toBeGreaterThan(group);
		expect(slots.layers[layerOf('strip6')]).toBeLessThan(group);

		const flat = morphOverlayElementSlots(plan);
		expect(flat.elements.get('band6')).toBeGreaterThan(flat.groups[0]!);
		expect(flat.elements.get('strip6')).toBeLessThan(flat.groups[0]!);
	});
});
