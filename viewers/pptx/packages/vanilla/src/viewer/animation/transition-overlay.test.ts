import type { PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { playTransitionOverlay } from './transition-overlay';

/**
 * What a viewer actually SEES during a slide change.
 *
 * Two regressions shipped here while every structural assertion still passed,
 * because the animations really were running: the layers existed, the keyframes
 * existed, and a human on a 1920x1080 display still saw a flicker and a cut.
 */

function buildStage(doc: Document, ids: string[]): HTMLElement {
	const stage = doc.createElement('div');
	stage.className = 'pptxv-stage';
	// What `renderSlideStage` paints: the resolved (always opaque) slide fill.
	stage.style.backgroundColor = 'rgb(255, 255, 255)';
	for (const id of ids) {
		const box = doc.createElement('div');
		box.dataset.elementId = id;
		stage.appendChild(box);
	}
	return stage;
}

/** One pairable shape per slide, so `buildMorphTransitionPlan` returns a plan. */
function morphSlide(id: string, x: number): PptxSlide {
	return {
		id,
		elements: [
			{
				id: `${id}-shape`,
				type: 'shape',
				name: 'Rectangle 1',
				x,
				y: 10,
				width: 100,
				height: 50,
				shapeType: 'rect',
			},
		],
	} as unknown as PptxSlide;
}

/** A full-slide picture whose frame never moves and whose source crop does. */
function rescaledPictureSlide(id: string, crop: Record<string, number>): PptxSlide {
	return {
		id,
		elements: [
			{
				id: `${id}-picture`,
				type: 'picture',
				name: '!!Background',
				x: 0,
				y: 0,
				width: 960,
				height: 540,
				imagePath: 'ppt/media/image9.png',
				...crop,
			},
		],
	} as unknown as PptxSlide;
}

/**
 * The shape of the issue #146 morph: an unchanged opaque disc, a backdrop that
 * departs (so the disc's ghost is kept as a shield), and new wording arriving
 * INSIDE the disc.
 */
function discAndWording(): { from: PptxSlide; to: PptxSlide } {
	const disc = {
		type: 'shape',
		name: '!!Content',
		x: 0,
		y: 0,
		width: 300,
		height: 300,
		shapeType: 'ellipse',
		shapeStyle: { fillMode: 'solid', fillColor: '#27282A' },
	};
	return {
		from: {
			id: 'out',
			elements: [
				{
					id: 'out-backdrop',
					type: 'shape',
					name: 'Backdrop',
					x: 0,
					y: 0,
					width: 960,
					height: 540,
				},
				{ ...disc, id: 'out-disc' },
			],
		} as unknown as PptxSlide,
		to: {
			id: 'in',
			elements: [
				{ ...disc, id: 'in-disc' },
				{
					id: 'in-wording',
					type: 'text',
					name: 'TextBox 9',
					x: 50,
					y: 60,
					width: 200,
					height: 30,
					text: 'Multi-Domain Fusion',
				},
			],
		} as unknown as PptxSlide,
	};
}

/**
 * Issue #141's slides 4 to 5: a departing picture keeps the unchanged
 * backdrop's ghost painted, a full-slide picture fades in above the backdrop,
 * and the title travels ABOVE that picture on the incoming slide.
 */
function veiledTitle(): { from: PptxSlide; to: PptxSlide } {
	const picture = (id: string, name: string, imagePath: string): unknown => ({
		id,
		name,
		type: 'picture',
		x: 0,
		y: 0,
		width: 1280,
		height: 720,
		imagePath,
		imageData: `data:${imagePath}`,
	});
	const title = (id: string, x: number): unknown => ({
		id,
		name: 'Title',
		type: 'text',
		x,
		y: 160,
		width: 490,
		height: 100,
		text: 'Our Services',
	});
	return {
		from: {
			id: 'out',
			elements: [
				picture('out-gone', 'Leaving', 'ppt/media/leaving.png'),
				picture('out-backdrop', 'Backdrop', 'ppt/media/backdrop.png'),
				title('out-title', 690),
			],
		} as unknown as PptxSlide,
		to: {
			id: 'in',
			elements: [
				picture('in-backdrop', 'Backdrop', 'ppt/media/backdrop.png'),
				picture('in-veil', 'Veil', 'ppt/media/veil.png'),
				title('in-title', 110),
			],
		} as unknown as PptxSlide,
	};
}

/**
 * The same panel with its wording REPLACED rather than arriving: a matched pair
 * the overlay paints both halves of, which is what has to be blended additively
 * (issue #161). `shapeId` is the identity that pairs two text boxes saying
 * different things - proximity alone deliberately refuses them (issue #131).
 */
function discAndReplacedWording(): { from: PptxSlide; to: PptxSlide } {
	const { from, to } = discAndWording();
	const wording = (id: string, text: string): unknown => ({
		id,
		type: 'text',
		name: 'TextBox 6',
		shapeId: 7,
		x: 50,
		y: 60,
		width: 200,
		height: 30,
		text,
	});
	return {
		from: {
			...from,
			elements: [...from.elements, wording('out-wording', 'Open Integration')],
		} as unknown as PptxSlide,
		to: {
			...to,
			elements: [
				...to.elements.filter((element) => element.id !== 'in-wording'),
				wording('in-wording', 'Tactical Edge'),
			],
		} as unknown as PptxSlide,
	};
}

describe('playTransitionOverlay', () => {
	let doc: Document;
	let stageWrap: HTMLElement;

	beforeEach(() => {
		vi.useFakeTimers();
		doc = document;
		stageWrap = doc.createElement('div');
		doc.body.appendChild(stageWrap);
	});

	afterEach(() => {
		vi.useRealTimers();
		stageWrap.remove();
	});

	function play(
		transition: PptxSlideTransition,
		morph = false,
	): { outgoing: HTMLElement; incoming: HTMLElement; cancel: () => void } {
		const outgoing = buildStage(doc, ['out-shape']);
		const incoming = buildStage(doc, ['in-shape']);
		const cancel = playTransitionOverlay({
			doc,
			stageWrap,
			outgoing,
			incoming,
			transition,
			outgoingSlide: morph ? morphSlide('out', 10) : undefined,
			incomingSlide: morph ? morphSlide('in', 400) : undefined,
			onDone: vi.fn(),
		});
		return { outgoing, incoming, cancel };
	}

	it('stretches each layer to the overlay instead of shrink-wrapping its content', () => {
		// The stage inside a layer scales with a CSS `transform`, which never
		// changes its laid-out box: an auto-sized `top`/`left` layer therefore
		// measures the deck's NATIVE size, and `overflow: hidden` then crops the
		// transition to that corner of a larger display (1280x720 of a 1920x1080
		// show). jsdom has no layout engine, so the declaration is asserted.
		const { cancel } = play({ type: 'fade', durationMs: 300 } as PptxSlideTransition);
		const layers = [...stageWrap.querySelectorAll<HTMLElement>('.pptxv-transition-layer')];
		expect(layers).toHaveLength(2);
		for (const layer of layers) {
			expect(layer.style.position).toBe('absolute');
			expect(layer.style.overflow).toBe('hidden');
			// jsdom keeps the authored `0`; a browser normalises it to `0px`. A
			// regression to `top`/`left` alone leaves `inset` empty and fails here.
			expect(['0', '0px'], 'a clipped layer must fill its overlay, not its content').toContain(
				layer.style.inset,
			);
		}
		cancel();
	});

	it('paints no slide background on the morph departing layer', () => {
		const { incoming, cancel } = play(
			{ type: 'morph', durationMs: 800 } as PptxSlideTransition,
			true,
		);
		const ghostStage = stageWrap.querySelector<HTMLElement>(
			'[data-pptx-morph-outgoing] .pptxv-stage',
		);
		expect(ghostStage).not.toBeNull();
		// The departing snapshot only carries the morphing shapes and sits ABOVE
		// the incoming slide, so it has to be see-through. Its own (always opaque)
		// slide background otherwise covered the whole morph with a flat slab.
		expect(ghostStage!.style.backgroundColor).toBe('transparent');
		// The incoming layer is a real surface and keeps its background.
		expect(incoming.style.backgroundColor).toBe('rgb(255, 255, 255)');
		cancel();
	});

	it('zooms a picture whose only change is its source crop', () => {
		// PowerPoint's "Scale Height"/"Scale Width" is an `a:srcRect` crop inside
		// an unchanged frame, so both halves of this pair agree on position, size,
		// blip and style: the engine saw an inert pair and the picture cut between
		// crops in a single frame (issue #148). The crop is painted on the `<img>`,
		// so the rule has to reach that node and not the element container.
		const cancel = playTransitionOverlay({
			doc,
			stageWrap,
			outgoing: buildStage(doc, ['out-picture']),
			incoming: buildStage(doc, ['in-picture']),
			transition: { type: 'morph', durationMs: 800 } as PptxSlideTransition,
			outgoingSlide: rescaledPictureSlide('out', { cropLeft: 0.05739, cropRight: 0.05739 }),
			incomingSlide: rescaledPictureSlide('in', { cropLeft: 0.00356, cropRight: 0.00356 }),
			onDone: vi.fn(),
		});

		const css = [...stageWrap.querySelectorAll('style')].map((node) => node.textContent).join('\n');
		expect(css).toContain('[data-element-id="in-picture"] img { animation: pptx-morph-crop-');
		expect(css).toContain('@keyframes pptx-morph-crop-');
		cancel();
	});

	// Regression (issue #146): the wheel deck's centre disc is identical on both
	// slides, so its opaque ghost sat over the wording dissolving in inside it -
	// invisible until the overlay came down. The plan hands those few over
	// separately and the overlay must paint them above every ghost.
	it('paints the arriving shapes the plan lifted, above the departing layer', () => {
		const outgoing = buildStage(doc, ['out-backdrop', 'out-disc']);
		const incoming = buildStage(doc, ['in-disc', 'in-wording']);
		const cancel = playTransitionOverlay({
			doc,
			stageWrap,
			outgoing,
			incoming,
			transition: { type: 'morph', durationMs: 800 } as PptxSlideTransition,
			outgoingSlide: discAndWording().from,
			incomingSlide: discAndWording().to,
			onDone: vi.fn(),
		});

		const lifted = stageWrap.querySelector<HTMLElement>('[data-pptx-morph-lifted]');
		expect(lifted).not.toBeNull();
		expect(lifted!.style.zIndex).toBe('3');
		// The clone is stripped to exactly what the plan lifted.
		expect(lifted!.querySelector('[data-element-id="in-wording"]')).not.toBeNull();
		expect(lifted!.querySelector('[data-element-id="in-disc"]')).toBeNull();

		const css = [...stageWrap.querySelectorAll('style')].map((node) => node.textContent).join('\n');
		expect(css).toContain('[data-pptx-morph-lifted] [data-element-id="in-wording"]');
		expect(css).toMatch(
			/\[data-pptx-morph-incoming\] \[data-element-id="in-wording"\] \{ animation: pptx-morph-lifted-hidden/u,
		);
		cancel();
	});

	// Regression (issue #141): every lifted arrival used to be painted above
	// every ghost, so a full-slide picture fading in veiled the title travelling
	// ABOVE it on the incoming slide, which then snapped bright at the end.
	it('stacks a travelling title ghost above a lifted picture arriving under it', () => {
		const { from, to } = veiledTitle();
		const cancel = playTransitionOverlay({
			doc,
			stageWrap,
			outgoing: buildStage(doc, ['out-gone', 'out-backdrop', 'out-title']),
			incoming: buildStage(doc, ['in-backdrop', 'in-veil', 'in-title']),
			transition: { type: 'morph', durationMs: 800 } as PptxSlideTransition,
			outgoingSlide: from,
			incomingSlide: to,
			onDone: vi.fn(),
		});

		const layers = [
			...stageWrap.querySelectorAll<HTMLElement>(
				'[data-pptx-morph-outgoing], [data-pptx-morph-lifted]',
			),
		].map((layer) => ({
			z: Number(layer.style.zIndex),
			ids: [...layer.querySelectorAll<HTMLElement>('[data-element-id]')].map(
				(node) => node.dataset.elementId,
			),
		}));
		const zOf = (id: string): number => layers.find((layer) => layer.ids.includes(id))!.z;
		expect(zOf('out-backdrop')).toBeLessThan(zOf('in-veil'));
		expect(zOf('in-veil')).toBeLessThan(zOf('out-title'));
		// The whole snapshot is never painted: only the runs the plan names.
		expect(stageWrap.querySelectorAll('[data-pptx-transition-layer="outgoing"]')).toHaveLength(2);
		cancel();
	});

	// Regression (issue #161): two fades stacked as ordinary layers composite
	// source-over, leaving the ink the halves SHARE at 0.75 of full strength
	// halfway through - measured 34.6/255 too dark on the wheel deck, which
	// bites chunks out of glyphs where the two line grids cross. PowerPoint's
	// own render keeps the blend coefficients summing to 1.0 for every frame, so
	// the pair is summed inside its own isolation group instead.
	it('paints a pair dissolving in place as one isolated, additive group', () => {
		const { from, to } = discAndReplacedWording();
		const cancel = playTransitionOverlay({
			doc,
			stageWrap,
			outgoing: buildStage(doc, ['out-backdrop', 'out-disc', 'out-wording']),
			incoming: buildStage(doc, ['in-disc', 'in-wording']),
			transition: { type: 'morph', durationMs: 800 } as PptxSlideTransition,
			outgoingSlide: from,
			incomingSlide: to,
			onDone: vi.fn(),
		});

		const group = stageWrap.querySelector<HTMLElement>('[data-pptx-morph-crossfade="in-wording"]');
		expect(group, 'the pair must be rendered as its own group').not.toBeNull();
		expect(group!.style.isolation).toBe('isolate');
		// Above the one ghost layer (2).
		expect(group!.style.zIndex).toBe('3');
		const halves = [...group!.querySelectorAll<HTMLElement>('.pptxv-transition-layer')];
		expect(halves).toHaveLength(2);
		for (const half of halves) {
			expect(half.style.mixBlendMode).toBe('plus-lighter');
			// The dissolve rides the wrapper: on the element it would take a
			// compositing layer whose raster snaps to whole device pixels and paints
			// the wording off the live stage.
			expect(half.style.animation).toContain('-fade');
		}
		// Each half is stripped to its own element...
		expect(halves[0].querySelector('[data-element-id="out-wording"]')).not.toBeNull();
		expect(halves[0].querySelector('[data-element-id="out-disc"]')).toBeNull();
		expect(halves[1].querySelector('[data-element-id="in-wording"]')).not.toBeNull();
		// ...and the ghost layer no longer paints the grouped one itself.
		const ghostLayer = stageWrap.querySelector<HTMLElement>(
			'[data-pptx-transition-layer="outgoing"][data-pptx-morph-outgoing]',
		);
		expect(ghostLayer!.querySelector('[data-element-id="out-wording"]')).toBeNull();
		expect(ghostLayer!.querySelector('[data-element-id="out-disc"]')).not.toBeNull();
		cancel();
	});

	// Regression: the plan names elements at whatever level the morph matched
	// them, so a ghost can be a whole GROUP (the wheel deck's centre panel
	// dissolves as one object). Sparing only the named node and its ancestors
	// stripped that group's children and the departing layer painted an empty
	// box, so the old panel simply never appeared.
	it('keeps the contents of a ghost that is itself a group', () => {
		// A group whose cast changed dissolves as ONE object, so the plan names the
		// group and the snapshot has to keep everything inside it.
		const panel = (id: string, children: unknown[]): unknown => ({
			id,
			type: 'group',
			name: '!!Circle',
			x: 505,
			y: 225,
			width: 270,
			height: 270,
			children,
		});
		const disc = (id: string): unknown => ({
			id,
			type: 'shape',
			name: '!!Content',
			x: 0,
			y: 0,
			width: 270,
			height: 270,
			shapeStyle: { fillMode: 'solid', fillColor: '#27282A' },
		});
		const outgoingSlide = {
			id: 'out',
			elements: [
				panel('out-group', [
					disc('out-disc'),
					{
						id: 'out-line',
						type: 'text',
						x: 28,
						y: 121,
						width: 214,
						height: 28,
						text: 'Select Challenge',
					},
				]),
			],
		} as unknown as PptxSlide;
		const incomingSlide = {
			id: 'in',
			elements: [
				panel('in-group', [
					disc('in-disc'),
					{
						id: 'in-a',
						type: 'text',
						x: 28,
						y: 61,
						width: 214,
						height: 29,
						text: 'Multi-Domain Fusion',
					},
					{
						id: 'in-b',
						type: 'text',
						x: 41,
						y: 95,
						width: 193,
						height: 36,
						text: 'Combining data',
					},
				]),
			],
		} as unknown as PptxSlide;

		const outgoing = doc.createElement('div');
		outgoing.className = 'pptxv-stage';
		const groupNode = doc.createElement('div');
		groupNode.dataset.elementId = 'out-group';
		const childNode = doc.createElement('div');
		childNode.dataset.elementId = 'out-disc';
		groupNode.appendChild(childNode);
		const strayNode = doc.createElement('div');
		strayNode.dataset.elementId = 'out-stray';
		outgoing.append(groupNode, strayNode);

		const cancel = playTransitionOverlay({
			doc,
			stageWrap,
			outgoing,
			incoming: buildStage(doc, ['in-group']),
			transition: { type: 'morph', durationMs: 800 } as PptxSlideTransition,
			outgoingSlide,
			incomingSlide,
			onDone: vi.fn(),
		});

		const ghosts = stageWrap.querySelector<HTMLElement>('[data-pptx-morph-outgoing]');
		expect(ghosts!.querySelector('[data-element-id="out-group"]')).not.toBeNull();
		expect(ghosts!.querySelector('[data-element-id="out-disc"]')).not.toBeNull();
		// Anything the plan did not name still goes.
		expect(ghosts!.querySelector('[data-element-id="out-stray"]')).toBeNull();
		cancel();
	});

	it('keeps the slide background on a non-morph transition snapshot', () => {
		const { outgoing, incoming, cancel } = play({
			type: 'fade',
			durationMs: 300,
		} as PptxSlideTransition);
		expect(outgoing.style.backgroundColor).toBe('rgb(255, 255, 255)');
		expect(incoming.style.backgroundColor).toBe('rgb(255, 255, 255)');
		cancel();
	});

	describe('multi-fragment cinematic transitions', () => {
		const FRAGMENT_TYPES = [
			'vortex',
			'honeycomb',
			'glitter',
			'shred',
			'fracture',
			'curtains',
			'airplane',
		] as const;

		it.each(FRAGMENT_TYPES)('renders N clipped fragments (not a single layer) for %s', (type) => {
			const { cancel } = play({ type, durationMs: 900 } as PptxSlideTransition);
			const fragments = [
				...stageWrap.querySelectorAll<HTMLElement>('[data-pptx-transition-fragment]'),
			];
			expect(fragments.length).toBeGreaterThan(1);
			expect(fragments.some((el) => el.style.clipPath.startsWith('polygon('))).toBeTruthy();
			expect(fragments.every((el) => el.style.willChange.includes('transform'))).toBeTruthy();
			// Each fragment carries its own clone of the stage's content.
			for (const fragment of fragments) {
				expect(fragment.querySelector('[data-element-id]')).not.toBeNull();
			}
			cancel();
		});

		it('a non-fragment type (fade) never emits fragment markers', () => {
			const { cancel } = play({ type: 'fade', durationMs: 300 } as PptxSlideTransition);
			expect(stageWrap.querySelectorAll('[data-pptx-transition-fragment]')).toHaveLength(0);
			cancel();
		});
	});
});

/** What a morph actually paints, read back from the overlay's DOM and `<style>` (issue #141). */
describe('playTransitionOverlay: what a morph paints (issue #141)', () => {
	let stageWrap: HTMLElement;
	/** Every overlay a test started, cancelled even when an assertion throws first. */
	const overlays: (() => void)[] = [];

	beforeEach(() => {
		vi.useFakeTimers();
		stageWrap = document.createElement('div');
		document.body.appendChild(stageWrap);
	});

	afterEach(() => {
		for (const cancel of overlays.splice(0)) {
			cancel();
		}
		vi.useRealTimers();
		stageWrap.remove();
	});

	function stageOf(elements: readonly { id: string; text?: string }[]): HTMLElement {
		const stage = buildStage(document, []);
		for (const element of elements) {
			const box = document.createElement('div');
			box.dataset.elementId = element.id;
			if (element.text !== undefined) {
				box.textContent = element.text;
			}
			stage.appendChild(box);
		}
		return stage;
	}

	function playMorph(
		from: PptxSlide,
		to: PptxSlide,
		outgoing: HTMLElement,
		incoming: HTMLElement,
	): () => void {
		const cancel = playTransitionOverlay({
			doc: document,
			stageWrap,
			outgoing,
			incoming,
			transition: { type: 'morph', durationMs: 1000 } as PptxSlideTransition,
			outgoingSlide: from,
			incomingSlide: to,
			onDone: vi.fn(),
		});
		overlays.push(cancel);
		return cancel;
	}

	const overlayCss = (): string =>
		[...stageWrap.querySelectorAll('style')].map((node) => node.textContent).join('\n');

	/** The opacity an animation shorthand holds at `progress` (0-100), from its keyframes. */
	function opacityAt(css: string, animation: string, progress: number): number {
		const name = animation.trim().split(/\s+/u)[0] ?? '';
		const start = css.indexOf(`@keyframes ${name} {`);
		if (start < 0) {
			return 1;
		}
		const block = css.slice(start, css.indexOf('\n}', start));
		const stops: { at: number; opacity: number }[] = [];
		for (const match of block.matchAll(
			/((?:from|to|[\d.]+%)(?:\s*,\s*(?:from|to|[\d.]+%))*)\s*\{([^}]*)\}/gu,
		)) {
			const opacity = /opacity:\s*([\d.]+)/u.exec(match[2] ?? '')?.[1];
			if (opacity === undefined) {
				continue;
			}
			for (const key of (match[1] ?? '').split(',')) {
				const stop = key.trim();
				const at = stop === 'from' ? 0 : stop === 'to' ? 100 : Number.parseFloat(stop);
				stops.push({ at, opacity: Number(opacity) });
			}
		}
		stops.sort((a, b) => a.at - b.at);
		const before = stops.filter((stop) => stop.at <= progress).at(-1);
		const after = stops.find((stop) => stop.at >= progress);
		if (!before || !after) {
			return (before ?? after)?.opacity ?? 1;
		}
		if (after.at === before.at) {
			return after.opacity;
		}
		return (
			before.opacity +
			((after.opacity - before.opacity) * (progress - before.at)) / (after.at - before.at)
		);
	}

	/**
	 * How opaque the overlay paints `node` at `progress`: the product of every
	 * animation reaching it, inline on it or an ancestor (a layer, a crossfade
	 * half) or through a scoped `[data-pptx-morph-*] [data-element-id]` rule.
	 */
	function paintedOpacity(node: HTMLElement, progress: number): number {
		const css = overlayCss();
		const id = node.dataset.elementId ?? '';
		let opacity = 1;
		for (let at: HTMLElement | null = node; at && at !== stageWrap; at = at.parentElement) {
			if (at.style.animation) {
				opacity *= opacityAt(css, at.style.animation, progress);
			}
			for (const scope of ['incoming', 'outgoing', 'lifted']) {
				if (!at.hasAttribute(`data-pptx-morph-${scope}`)) {
					continue;
				}
				const rule = css
					.split('\n')
					.find((line) =>
						line.startsWith(`[data-pptx-morph-${scope}] [data-element-id="${id}"] { animation:`),
					);
				const animation = rule ? /animation: ([^;]+);/u.exec(rule)?.[1] : undefined;
				if (animation !== undefined) {
					opacity *= opacityAt(css, animation, progress);
				}
			}
		}
		return opacity;
	}

	// Issue #141, slides 6 -> 7: the "L" column changes font, so it dissolves as
	// a crossfade group, while it slides under the top band the incoming slide
	// stacks ABOVE it. Painted above every overlay layer, the column showed
	// across the band for the whole morph and vanished from it in one frame.
	it('stacks a dissolving column between the band above it and the strip below it', () => {
		const picture = (id: string, name: string, y: number, height: number) => ({
			id,
			name,
			type: 'picture',
			x: 0,
			y,
			width: 1280,
			height,
			imagePath: `ppt/media/${name}.png`,
		});
		const column = (id: string, y: number, fontFamily: string) => ({
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
		});
		const from = {
			id: '6',
			slideNumber: 1,
			elements: [
				picture('bg6', 'backdrop', 0, 720),
				column('col6', -916, 'Calibri Light'),
				picture('strip6', 'strip', -1, 720),
				picture('band6', 'band', -1, 196),
			],
		} as unknown as PptxSlide;
		const to = {
			id: '7',
			slideNumber: 2,
			elements: [
				picture('bg7', 'backdrop', 0, 720),
				picture('strip7', 'strip', -3, 720),
				column('col7', -158, 'Calibri'),
				picture('band7', 'band', -2, 196),
			],
		} as unknown as PptxSlide;
		const cancel = playMorph(from, to, stageOf(from.elements), stageOf(to.elements));

		const group = stageWrap.querySelector<HTMLElement>('[data-pptx-morph-crossfade="col7"]');
		expect(group, 'the column must dissolve as a group').not.toBeNull();
		const layerZ = (id: string): number => {
			const layer = [
				...stageWrap.querySelectorAll<HTMLElement>(
					'[data-pptx-transition-layer="outgoing"][data-pptx-morph-outgoing], [data-pptx-transition-layer="lifted"][data-pptx-morph-lifted]',
				),
			].find(
				(candidate) =>
					!candidate.closest('[data-pptx-morph-crossfade]') &&
					candidate.querySelector(`[data-element-id="${id}"]`) !== null,
			);
			expect(layer, `${id} must be painted in an overlay layer`).toBeDefined();
			return Number(layer!.style.zIndex);
		};
		const groupZ = Number(group!.style.zIndex);
		expect(layerZ('band6')).toBeGreaterThan(groupZ);
		expect(layerZ('strip6')).toBeLessThan(groupZ);
		cancel();
	});

	// Issue #141, slides 17 -> 18 and back: a three-paragraph text box parked off
	// the left edge whose second paragraph changes size flies onto the slide (or
	// back off it). Cross-dissolving the two layouts painted the paragraph twice
	// for the whole flight; one travelling copy must be the only one seen.
	describe('paints a size-changed paragraph flying from off-stage exactly once', () => {
		const TEXT = 'Parked opener\nSecond paragraph\nClosing line';
		const box = (id: string, x: number, fontSize?: number) => ({
			id,
			type: 'text',
			name: 'TextBox 9',
			x,
			y: 120,
			width: 300,
			height: 240,
			text: TEXT,
			textSegments: [
				{ text: 'Parked opener', style: { fontFamily: 'Arial', fontSize: 18 } },
				{ text: '\n' },
				{
					text: 'Second paragraph',
					style: { fontFamily: 'Arial', ...(fontSize === undefined ? {} : { fontSize }) },
				},
				{ text: '\n' },
				{ text: 'Closing line', style: { fontFamily: 'Arial', fontSize: 18 } },
			],
		});
		// Only on the source slide, so the travelling copy is painted as a ghost.
		const panel = (id: string) => ({
			id,
			type: 'shape',
			name: 'Panel',
			x: -600,
			y: 0,
			width: 1200,
			height: 720,
			shapeStyle: { fillMode: 'solid', fillColor: '#123456' },
		});
		const slideOf = (id: string, elements: unknown[]): PptxSlide =>
			({ id, slideNumber: 1, elements }) as unknown as PptxSlide;

		it.each([
			[
				'forward, 17 -> 18',
				slideOf('17', [panel('p17'), box('t17', -400, 14)]),
				slideOf('18', [box('t18', 875)]),
			],
			[
				'backward, 18 -> 17',
				slideOf('18', [panel('p18'), box('t18', 875)]),
				slideOf('17', [box('t17', -400, 14)]),
			],
		])('%s', (_, from, to) => {
			const withText = (slide: PptxSlide) =>
				slide.elements.map((element) => ({
					id: element.id,
					...(element.type === 'text' ? { text: TEXT } : {}),
				}));
			const cancel = playMorph(from, to, stageOf(withText(from)), stageOf(withText(to)));

			const copies = [...stageWrap.querySelectorAll<HTMLElement>('[data-element-id]')].filter(
				(node) => node.textContent === TEXT,
			);
			// The travelling ghost and the incoming layer's own copy are both in the DOM...
			expect(copies.length).toBeGreaterThanOrEqual(2);
			expect(
				copies.some((node) => node.closest('[data-pptx-morph-outgoing]') !== null),
				'the ghost carries the text',
			).toBeTruthy();
			// ...but only one of them can be seen halfway through the flight.
			const visible = copies.filter((node) => paintedOpacity(node, 50) > 0.01);
			expect(visible).toHaveLength(1);
			cancel();
		});
	});
});
