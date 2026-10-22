import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	buildMorphIncomingAnchoredOrder,
	ghostMorphPairsHiddenByGhosts,
	resolveMorphOverlayLayers,
} from './morph-overlay-layers';
import { buildMorphTransitionPlan } from './morph-plan';
import type { MorphPair } from './morph-types';

function box(id: string): PptxElement {
	return { id, name: id, type: 'shape', x: 0, y: 0, width: 100, height: 50 } as PptxElement;
}

function pair(from: PptxElement, to: PptxElement): MorphPair {
	return { fromElement: from, toElement: to };
}

const ids = (layers: ReturnType<typeof resolveMorphOverlayLayers>): string[][] =>
	layers.map((layer) => [layer.side, ...layer.elements.map((element) => element.id)]);

/**
 * The shape of issue #141's corporate deck, slides 4 to 5: a departing
 * full-slide picture keeps the unchanged backdrop's ghost painted, a different
 * full-slide picture fades in above the backdrop, and the title travels
 * horizontally ABOVE that picture on the incoming slide.
 */
function veiledTitle(): { from: PptxSlide; to: PptxSlide } {
	const picture = (id: string, name: string, imagePath: string): PptxElement =>
		({
			id,
			name,
			type: 'picture',
			x: 0,
			y: 0,
			width: 1280,
			height: 720,
			imagePath,
			imageData: `data:${imagePath}`,
		}) as unknown as PptxElement;
	const title = (id: string, x: number): PptxElement =>
		({
			id,
			name: 'Title',
			type: 'text',
			x,
			y: 160,
			width: 490,
			height: 100,
			text: 'Our Services',
		}) as unknown as PptxElement;
	return {
		from: {
			id: 'out',
			elements: [
				picture('out-gone', 'Leaving', 'ppt/media/leaving.png'),
				picture('out-backdrop', 'Backdrop', 'ppt/media/backdrop.png'),
				title('out-title', 690),
			],
		} as PptxSlide,
		to: {
			id: 'in',
			elements: [
				picture('in-backdrop', 'Backdrop', 'ppt/media/backdrop.png'),
				picture('in-veil', 'Veil', 'ppt/media/veil.png'),
				title('in-title', 110),
			],
		} as PptxSlide,
	};
}

describe('buildMorphIncomingAnchoredOrder', () => {
	it('ranks a pair where the incoming slide stacks it', () => {
		// Outgoing: title, photo. Incoming: photo, title. The photo slides in over
		// the title on the outgoing slide's order, under it on the incoming one.
		const outgoing = [box('a-title'), box('a-photo')];
		const incoming = [box('b-photo'), box('b-title')];
		const rank = buildMorphIncomingAnchoredOrder(outgoing, incoming, [
			pair(outgoing[0], incoming[1]),
			pair(outgoing[1], incoming[0]),
		]);

		expect(rank.get('a-photo')!).toBeLessThan(rank.get('a-title')!);
		expect(rank.get('a-title')).toBe(rank.get('b-title'));
	});

	it('keeps a departure between the shapes that surrounded it', () => {
		const outgoing = [box('a-base'), box('a-leaving'), box('a-cap')];
		const incoming = [box('b-base'), box('b-cap')];
		const rank = buildMorphIncomingAnchoredOrder(outgoing, incoming, [
			pair(outgoing[0], incoming[0]),
			pair(outgoing[2], incoming[1]),
		]);

		expect(rank.get('a-base')!).toBeLessThan(rank.get('a-leaving')!);
		expect(rank.get('a-leaving')!).toBeLessThan(rank.get('a-cap')!);
	});

	it('keeps a departure as low as its neighbours allow', () => {
		// The departing background sits at the bottom of the outgoing slide; the
		// title above it jumps to the top of the incoming one. The background must
		// not ride up with it over the arrivals.
		const outgoing = [box('a-bg'), box('a-title')];
		const incoming = [box('b-new'), box('b-photo'), box('b-title')];
		const rank = buildMorphIncomingAnchoredOrder(outgoing, incoming, [
			pair(outgoing[1], incoming[2]),
		]);

		expect(rank.get('a-bg')!).toBeLessThan(rank.get('b-new')!);
	});

	it('keeps a departure under the shape that covered it when its neighbours cross', () => {
		// Issue #141's slides 9 to 10: a column at the bottom of the outgoing slide
		// climbs to the top of the incoming one; the departing picture above it was
		// hidden by the backdrop and has to stay under the backdrop.
		const outgoing = [box('a-column'), box('a-leaving'), box('a-backdrop')];
		const incoming = [box('b-backdrop'), box('b-blob'), box('b-column')];
		const rank = buildMorphIncomingAnchoredOrder(outgoing, incoming, [
			pair(outgoing[0], incoming[2]),
			pair(outgoing[2], incoming[0]),
		]);

		expect(rank.get('a-leaving')!).toBeLessThan(rank.get('a-backdrop')!);
		expect(rank.get('a-backdrop')!).toBeLessThan(rank.get('a-column')!);
	});
});

describe('ghostMorphPairsHiddenByGhosts', () => {
	const tall = (id: string, y: number): PptxElement =>
		({ id, name: id, type: 'shape', x: 0, y, width: 300, height: 700 }) as PptxElement;

	it('ghosts a pair a holding ghost would hide although it belongs above it', () => {
		const outgoing = [tall('a-column', 760), tall('a-backdrop', 0)];
		const incoming = [tall('b-backdrop', 0), tall('b-column', 0)];
		const pairs = [pair(outgoing[0], incoming[1]), pair(outgoing[1], incoming[0])];

		const ghosts = ghostMorphPairsHiddenByGhosts(
			new Set(['a-backdrop']),
			outgoing,
			incoming,
			pairs,
			() => true,
		);

		expect([...ghosts]).toStrictEqual(['a-backdrop', 'a-column']);
	});

	it('leaves a pair that belongs under the ghost, or one a dissolving ghost covers', () => {
		const outgoing = [tall('a-column', 760), tall('a-backdrop', 0)];
		const below = [tall('b-column', 0), tall('b-backdrop', 0)];
		const pairs = [pair(outgoing[0], below[0]), pair(outgoing[1], below[1])];
		expect([
			...ghostMorphPairsHiddenByGhosts(new Set(['a-backdrop']), outgoing, below, pairs, () => true),
		]).toStrictEqual(['a-backdrop']);

		const above = [tall('b-backdrop', 0), tall('b-column', 0)];
		const swapped = [pair(outgoing[0], above[1]), pair(outgoing[1], above[0])];
		expect([
			...ghostMorphPairsHiddenByGhosts(
				new Set(['a-backdrop']),
				outgoing,
				above,
				swapped,
				() => false,
			),
		]).toStrictEqual(['a-backdrop']);
	});
});

describe('resolveMorphOverlayLayers', () => {
	it('interleaves ghosts and lifted arrivals in one order', () => {
		const outgoing = [box('a-backdrop'), box('a-title')];
		const incoming = [box('b-backdrop'), box('b-veil'), box('b-title')];
		const pairs = [pair(outgoing[0], incoming[0]), pair(outgoing[1], incoming[2])];

		const layers = resolveMorphOverlayLayers(outgoing, [incoming[1]], outgoing, incoming, pairs);

		expect(ids(layers)).toStrictEqual([
			['outgoing', 'a-backdrop'],
			['incoming', 'b-veil'],
			['outgoing', 'a-title'],
		]);
	});

	it('paints a lifted half over its own ghost', () => {
		const outgoing = [box('a-wording')];
		const incoming = [box('b-wording')];
		const layers = resolveMorphOverlayLayers(outgoing, incoming, outgoing, incoming, [
			pair(outgoing[0], incoming[0]),
		]);

		expect(ids(layers)).toStrictEqual([
			['outgoing', 'a-wording'],
			['incoming', 'b-wording'],
		]);
	});

	it('is empty when the overlay paints nothing', () => {
		expect(resolveMorphOverlayLayers([], [], [box('a')], [box('b')], [])).toStrictEqual([]);
	});
});

describe('buildMorphTransitionPlan overlayLayers', () => {
	// Regression (issue #141): every lifted arrival used to be painted above
	// every ghost, so the full-slide picture fading in veiled the travelling
	// title at half strength for the whole morph, and the title snapped bright
	// the frame the overlay came down.
	it('keeps a travelling title above a full-slide picture arriving under it', () => {
		const { from, to } = veiledTitle();
		const plan = buildMorphTransitionPlan(from, to, 2000)!;

		expect(plan.overlayIncomingElements.map((element) => element.id)).toStrictEqual(['in-veil']);
		expect(ids(plan.overlayLayers)).toStrictEqual([
			['outgoing', 'out-gone', 'out-backdrop'],
			['incoming', 'in-veil'],
			['outgoing', 'out-title'],
		]);
	});

	// Regression (issue #141, slides 9 to 10): the column picture is first in
	// the outgoing slide's order, so nothing painted sat below it and it was left
	// to climb on the live stage, under the overlay's opaque backdrop ghost. It
	// stayed invisible for the whole morph and popped in when the overlay went.
	it('paints a climbing column above the backdrop ghost that would hide it', () => {
		const picture = (id: string, name: string, y: number, width: number): PptxElement =>
			({
				id,
				name,
				type: 'picture',
				x: 0,
				y,
				width,
				height: 720,
				imagePath: `ppt/media/${name}.png`,
				imageData: `data:${name}`,
			}) as unknown as PptxElement;
		const from = {
			id: 'out',
			elements: [
				picture('out-column', 'Column', 760, 320),
				picture('out-gone', 'Leaving', 0, 1280),
				picture('out-backdrop', 'Backdrop', 0, 1280),
			],
		} as PptxSlide;
		const to = {
			id: 'in',
			elements: [
				picture('in-backdrop', 'Backdrop', 0, 1280),
				picture('in-column', 'Column', 0, 320),
			],
		} as PptxSlide;

		const plan = buildMorphTransitionPlan(from, to, 2000)!;

		expect(ids(plan.overlayLayers)).toStrictEqual([
			['outgoing', 'out-gone', 'out-backdrop', 'out-column'],
		]);
		expect(plan.outgoingAnimations.get('out-column')).toContain('pptx-morph-ghost-');
	});

	it('covers exactly the ghosts and lifted arrivals', () => {
		const { from, to } = veiledTitle();
		const plan = buildMorphTransitionPlan(from, to, 2000)!;
		const painted = plan.overlayLayers.flatMap((layer) =>
			layer.elements.map((element) => element.id),
		);

		expect(painted.toSorted()).toStrictEqual(
			[...plan.outgoingElements, ...plan.overlayIncomingElements]
				.map((element) => element.id)
				.toSorted(),
		);
	});
});
