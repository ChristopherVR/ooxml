import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { playAnimationSound, stopAnimationSound } from './animation-sound';
import PresentationTransitionOverlay from './PresentationTransitionOverlay.svelte';

vi.mock(import('./animation-sound'), () => ({
	playAnimationSound: vi.fn(),
	stopAnimationSound: vi.fn(),
}));

/**
 * What a viewer actually SEES during a slide change.
 *
 * Two regressions shipped here while every structural assertion in the suite
 * still passed, because the animations really were running:
 * `document.getAnimations()` reported dozens of them while a human on a
 * 1920x1080 display saw a flicker and a hard cut.
 */

/** The component's own text; jsdom gives no layout, so the CSS rule is read. */
const overlaySource = readFileSync(
	[
		'src/viewer/presentation/PresentationTransitionOverlay.svelte',
		'packages/svelte/src/viewer/presentation/PresentationTransitionOverlay.svelte',
	]
		.map((candidate) => resolve(process.cwd(), candidate))
		.find((candidate) => existsSync(candidate))!,
	'utf8',
);

const canvasSize = { width: 960, height: 540 };

function makeSlide(id: string, elements: unknown[] = []): PptxSlide {
	return { id, elements, backgroundColor: '#ffffff' } as unknown as PptxSlide;
}

/**
 * One shape per slide, so `buildMorphTransitionPlan` has something to pair and
 * actually returns a plan (it bails out on two empty slides).
 */
function morphable(slideId: string, x: number): PptxSlide {
	return makeSlide(slideId, [
		{
			id: `${slideId}-shape`,
			type: 'shape',
			name: 'Rectangle 1',
			x,
			y: 10,
			width: 100,
			height: 50,
			shapeType: 'rect',
		},
	]);
}

/** A full-slide picture whose frame never moves and whose source crop does. */
function rescaledPicture(slideId: string, crop: Record<string, number>): PptxSlide {
	return makeSlide(slideId, [
		{
			id: `${slideId}-picture`,
			type: 'picture',
			name: '!!Background',
			x: 0,
			y: 0,
			width: 960,
			height: 540,
			imagePath: 'ppt/media/image9.png',
			...crop,
		},
	]);
}

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

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
		from: makeSlide('out', [
			{ id: 'out-backdrop', type: 'shape', name: 'Backdrop', x: 0, y: 0, width: 960, height: 540 },
			{ ...disc, id: 'out-disc' },
		]),
		to: makeSlide('in', [
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
		]),
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
		width: 960,
		height: 540,
		imagePath,
		imageData: `data:${imagePath}`,
	});
	const title = (id: string, x: number): unknown => ({
		id,
		name: 'Title',
		type: 'text',
		x,
		y: 160,
		width: 400,
		height: 100,
		text: 'Our Services',
	});
	return {
		from: makeSlide('out', [
			picture('out-gone', 'Leaving', 'ppt/media/leaving.png'),
			picture('out-backdrop', 'Backdrop', 'ppt/media/backdrop.png'),
			title('out-title', 520),
		]),
		to: makeSlide('in', [
			picture('in-backdrop', 'Backdrop', 'ppt/media/backdrop.png'),
			picture('in-veil', 'Veil', 'ppt/media/veil.png'),
			title('in-title', 80),
		]),
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
		from: makeSlide('out', [
			...(from.elements as unknown[]),
			wording('out-wording', 'Open Integration'),
		]),
		to: makeSlide('in', [
			...(to.elements as unknown[]).filter(
				(element) => (element as { id: string }).id !== 'in-wording',
			),
			wording('in-wording', 'Tactical Edge'),
		]),
	};
}

function mountOverlay(
	transition: PptxSlideTransition,
	slides?: { from: PptxSlide; to: PptxSlide },
): HTMLElement {
	const morph = transition.type === 'morph';
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(PresentationTransitionOverlay, {
		target,
		props: {
			outgoingSlide: slides?.from ?? (morph ? morphable('out', 10) : makeSlide('out')),
			incomingSlide: slides?.to ?? (morph ? morphable('in', 400) : makeSlide('in')),
			canvasSize,
			mediaDataUrls: new Map<string, string>(),
			transition,
			ondone: vi.fn(),
		},
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

describe('presentationTransitionOverlay', () => {
	it('paints no slide background on the morph departing layer', () => {
		const target = mountOverlay({ type: 'morph', durationMs: 800 } as PptxSlideTransition);
		expect(target.querySelectorAll('[data-pptx-morph-outgoing]')).not.toHaveLength(0);

		expect(target.querySelectorAll('.pptx-svelte-stage')).toHaveLength(2);
		const outgoing = target.querySelector<HTMLElement>(
			'[data-pptx-morph-outgoing] .pptx-svelte-stage',
		)!;
		const incoming = target.querySelector<HTMLElement>(
			'[data-pptx-morph-incoming] .pptx-svelte-stage',
		)!;
		// The departing shapes float over the incoming slide, so this layer has to
		// be see-through. `getSlideBackgroundStyle` always resolves to an OPAQUE
		// paint, so inheriting it hid the whole morph behind a flat slab for the
		// entire transition.
		expect(outgoing.getAttribute('style')).toMatch(/transparent/u);
		expect(outgoing.getAttribute('style')).not.toContain('#ffffff');
		// The incoming layer is a real surface and keeps its background.
		expect(incoming.getAttribute('style')).toContain('#ffffff');
	});

	it('zooms a picture whose only change is its source crop', () => {
		// PowerPoint's "Scale Height"/"Scale Width" is an `a:srcRect` crop inside
		// an unchanged frame, so both halves of this pair agree on position, size,
		// blip and style: the engine saw an inert pair and the picture cut between
		// crops in a single frame (issue #148). The crop is painted on the `<img>`,
		// so the rule has to reach that node and not the element container.
		const target = document.createElement('div');
		document.body.appendChild(target);
		const instance = mount(PresentationTransitionOverlay, {
			target,
			props: {
				outgoingSlide: rescaledPicture('out', { cropLeft: 0.05739, cropRight: 0.05739 }),
				incomingSlide: rescaledPicture('in', { cropLeft: 0.00356, cropRight: 0.00356 }),
				canvasSize,
				mediaDataUrls: new Map<string, string>(),
				transition: { type: 'morph', durationMs: 800 } as PptxSlideTransition,
				ondone: vi.fn(),
			},
		});
		flushSync();
		cleanup = () => {
			unmount(instance);
			target.remove();
		};

		const css = [...target.querySelectorAll('style')].map((node) => node.textContent).join('\n');
		expect(css).toContain('[data-element-id="in-picture"] img { animation: pptx-morph-crop-');
		expect(css).toContain('@keyframes pptx-morph-crop-');
	});

	it('sees a turning flip from the centre of its own canvas', () => {
		// Shared `morph-flip-turn`: the eye is (width / 2) / tan(30 deg) away,
		// 831.384 px for this 960 px canvas (the 16:9 default is 1108.513 px).
		const tri = (slideId: string, flipHorizontal: boolean) =>
			makeSlide(slideId, [
				{
					id: `${slideId}-tri`,
					type: 'shape',
					name: '!!Tri',
					x: 100,
					y: 100,
					width: 300,
					height: 200,
					shapeType: 'rtTriangle',
					flipHorizontal,
				},
			]);
		const target = mountOverlay({ type: 'morph', durationMs: 800 } as PptxSlideTransition, {
			from: tri('out', false),
			to: tri('in', true),
		});
		const css = [...target.querySelectorAll('style')].map((node) => node.textContent).join('\n');
		expect(css).toContain('perspective(831.384px)');
	});

	it('keeps the slide background on a non-morph transition snapshot', () => {
		const target = mountOverlay({ type: 'fade', durationMs: 300 } as PptxSlideTransition);
		const stages = [...target.querySelectorAll<HTMLElement>('.pptx-svelte-stage')];
		expect(stages).toHaveLength(2);
		for (const stage of stages) {
			expect(stage.getAttribute('style')).toContain('#ffffff');
		}
	});

	it('renders the morph stylesheet as text, so a hostile element id injects no markup', () => {
		// Element ids come from slide part paths. A closing `</style>` plus an
		// `<img onerror>` in the id must stay inert text, not become DOM.
		const evil = '</style><img src=x onerror=alert(1)>';
		const shape = (id: string, x: number) => ({
			id,
			type: 'shape',
			name: 'Rectangle 1',
			x,
			y: 10,
			width: 100,
			height: 50,
			shapeType: 'rect',
		});
		const target = mountOverlay({ type: 'morph', durationMs: 800 } as PptxSlideTransition, {
			from: makeSlide('out', [shape(evil, 10)]),
			to: makeSlide('in', [shape(evil, 400)]),
		});

		expect(target.querySelector('img')).toBeNull();
		const styles = [...target.querySelectorAll('style')];
		const css = styles.map((node) => node.textContent).join('\n');
		expect(css).toContain('[data-element-id="\\3c /style\\3e');
		expect(css).not.toContain('<img');
	});

	// Regression (issue #146): the wheel deck's centre disc is identical on both
	// slides, so its opaque ghost sat over the title, body and button dissolving
	// in inside it - they were invisible until the overlay came down. The plan
	// hands those few over separately and the overlay must paint them, above
	// every ghost, from the INCOMING slide.
	it('paints the arriving shapes the plan lifted, above the departing layer', () => {
		const target = mountOverlay(
			{ type: 'morph', durationMs: 800 } as PptxSlideTransition,
			discAndWording(),
		);

		const lifted = target.querySelector<HTMLElement>('[data-pptx-morph-lifted]');
		expect(lifted).not.toBeNull();
		expect(lifted!.getAttribute('style')).toContain('z-index: 3');
		expect(lifted!.textContent).toContain('Multi-Domain Fusion');

		const css = [...target.querySelectorAll('style')].map((node) => node.textContent).join('\n');
		expect(css).toContain('[data-pptx-morph-lifted] [data-element-id="in-wording"]');
		// Its copy on the incoming layer is held invisible, so the two never
		// composite with each other.
		expect(css).toMatch(
			/\[data-pptx-morph-incoming\] \[data-element-id="in-wording"\] \{ animation: pptx-morph-lifted-hidden/u,
		);
	});

	// Regression (issue #141): every lifted arrival used to be painted above
	// every ghost, so a full-slide picture fading in veiled the title travelling
	// ABOVE it on the incoming slide, which then snapped bright at the end.
	it('stacks a travelling title ghost above a lifted picture arriving under it', () => {
		const target = mountOverlay(
			{ type: 'morph', durationMs: 800 } as PptxSlideTransition,
			veiledTitle(),
		);

		const layers = [
			...target.querySelectorAll<HTMLElement>(
				'[data-pptx-morph-outgoing], [data-pptx-morph-lifted]',
			),
		].map((layer) => ({
			z: Number(/z-index:\s*(\d+)/u.exec(layer.getAttribute('style') ?? '')?.[1]),
			ids: [...layer.querySelectorAll<HTMLElement>('[data-element-id]')].map(
				(node) => node.dataset.elementId,
			),
		}));
		const zOf = (id: string): number => layers.find((layer) => layer.ids.includes(id))!.z;
		expect(zOf('out-backdrop')).toBeLessThan(zOf('in-veil'));
		expect(zOf('in-veil')).toBeLessThan(zOf('out-title'));
	});

	// Regression (issue #161): two fades stacked as ordinary layers composite
	// source-over, leaving the ink the halves SHARE at 0.75 of full strength
	// halfway through - measured 34.6/255 too dark on the wheel deck, which
	// bites chunks out of glyphs where the two line grids cross. PowerPoint's
	// own render keeps the blend coefficients summing to 1.0 for every frame, so
	// the pair is summed inside its own isolation group instead.
	it('paints a pair dissolving in place as one isolated, additive group', () => {
		const target = mountOverlay(
			{ type: 'morph', durationMs: 800 } as PptxSlideTransition,
			discAndReplacedWording(),
		);

		const group = target.querySelector<HTMLElement>('[data-pptx-morph-crossfade="in-wording"]');
		expect(group, 'the pair must be rendered as its own group').not.toBeNull();
		expect(group!.getAttribute('style')).toContain('isolation: isolate');
		// Above the one ghost layer (2).
		expect(group!.getAttribute('style')).toContain('z-index: 3');
		const halves = [...group!.querySelectorAll<HTMLElement>('[data-pptx-transition-layer]')];
		expect(halves).toHaveLength(2);
		for (const half of halves) {
			expect(half.getAttribute('style')).toContain('mix-blend-mode: plus-lighter');
			// The dissolve rides the wrapper: on the element it would take a
			// compositing layer whose raster snaps to whole device pixels and paints
			// the wording off the live stage.
			expect(half.getAttribute('style')).toContain('-fade');
		}
		expect(group!.textContent).toContain('Open Integration');
		expect(group!.textContent).toContain('Tactical Edge');
		// Painted exactly once: not left in the flat layers as well.
		expect(target.textContent!.match(/Open Integration/gu)).toHaveLength(1);
	});

	it('stretches each transition layer to the overlay instead of shrink-wrapping it', () => {
		// jsdom has no layout engine, so this is asserted on the rule itself. The
		// stage inside a layer scales with `transform`, which never changes its
		// laid-out box: an auto-sized `top`/`left` layer therefore measures the
		// deck's NATIVE size, and `overflow: hidden` then crops the transition to
		// that corner of a larger display (1280x720 of a 1920x1080 show).
		const layerRule =
			/\.pptx-svelte-transition-layer\s*\{([^}]*)\}/u.exec(overlaySource)?.[1] ?? '';
		expect(layerRule).toContain('position: absolute');
		expect(layerRule).toContain('overflow: hidden');
		expect(layerRule, 'a clipped layer must fill its overlay, not its content').toMatch(
			/inset:\s*0|width:\s*100%/u,
		);
		expect(layerRule).not.toMatch(/top:\s*0;\s*\n?\s*left:\s*0;/u);
	});
});

describe('presentationTransitionOverlay sound (p:sndAc/p:stSnd, p:endSnd)', () => {
	afterEach(() => {
		vi.mocked(playAnimationSound).mockClear();
		vi.mocked(stopAnimationSound).mockClear();
	});

	function mountWithTransition(
		transition: PptxSlideTransition,
		mediaDataUrls: Map<string, string>,
	) {
		const target = document.createElement('div');
		document.body.appendChild(target);
		const instance = mount(PresentationTransitionOverlay, {
			target,
			props: {
				outgoingSlide: makeSlide('out'),
				incomingSlide: makeSlide('in'),
				canvasSize,
				mediaDataUrls,
				transition,
				ondone: vi.fn(),
			},
		});
		flushSync();
		cleanup = () => {
			unmount(instance);
			target.remove();
		};
	}

	it('resolves the transition sound path through mediaDataUrls and plays it', () => {
		mountWithTransition(
			{
				type: 'fade',
				durationMs: 300,
				soundPath: 'ppt/media/media3.wav',
				soundLoop: true,
			} as PptxSlideTransition,
			new Map([['ppt/media/media3.wav', 'blob:sound']]),
		);
		expect(playAnimationSound).toHaveBeenCalledWith('blob:sound', true);
	});

	it('does not play when the raw archive path has no resolved URL yet', () => {
		mountWithTransition(
			{ type: 'fade', durationMs: 300, soundPath: 'ppt/media/media3.wav' } as PptxSlideTransition,
			new Map(),
		);
		expect(playAnimationSound).not.toHaveBeenCalled();
	});

	it('stops the current sound for p:endSndAc (transition.stopSound)', () => {
		mountWithTransition(
			{ type: 'fade', durationMs: 300, stopSound: true } as PptxSlideTransition,
			new Map(),
		);
		expect(stopAnimationSound).toHaveBeenCalledOnce();
		expect(playAnimationSound).not.toHaveBeenCalled();
	});

	it('does NOT stop the sound merely because the overlay unmounts (Loop Until Next Sound)', () => {
		mountWithTransition(
			{ type: 'fade', durationMs: 300, soundPath: 'a.wav', soundLoop: true } as PptxSlideTransition,
			new Map([['a.wav', 'blob:sound']]),
		);
		vi.mocked(stopAnimationSound).mockClear();
		cleanup?.();
		cleanup = undefined;
		expect(stopAnimationSound).not.toHaveBeenCalled();
	});
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
		const target = mountOverlay({ type, durationMs: 900 } as PptxSlideTransition);
		const fragments = target.querySelectorAll('[data-pptx-transition-fragment]');
		expect(fragments.length).toBeGreaterThan(1);
		const clipPaths = [...fragments].map((el) => (el as HTMLElement).style.clipPath);
		expect(clipPaths.some((clipPath) => clipPath.startsWith('polygon('))).toBeTruthy();
		const willChange = [...fragments].map((el) => (el as HTMLElement).style.willChange);
		expect(willChange.every((value) => value.includes('transform'))).toBeTruthy();
	});

	it('a non-fragment type (fade) never emits fragment markers', () => {
		const target = mountOverlay({ type: 'fade', durationMs: 300 } as PptxSlideTransition);
		expect(target.querySelectorAll('[data-pptx-transition-fragment]')).toHaveLength(0);
	});
});

const PIXEL = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';

function bandPicture(id: string, name: string, y: number, height: number): unknown {
	return {
		id,
		name,
		type: 'picture',
		x: 0,
		y,
		width: 1280,
		height,
		imagePath: `ppt/media/${name}.png`,
		imageData: PIXEL,
	};
}

/** One of issue #141's "L" columns: the same runs, another font, another place. */
function letterColumn(id: string, y: number, fontFamily: string): unknown {
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
	};
}

// Regression (issue #141, slides 6 -> 7): a column of "L"s changes font, so it
// dissolves as a crossfade group, while it slides under the top band the
// incoming slide stacks ABOVE it. Every group used to be painted above every
// overlay layer, so the column showed across the band for the whole morph and
// vanished from it in one frame when the overlay came down.
describe('morph crossfade groups in the overlay stack', () => {
	it('stacks the band above the dissolving column and the strip below it', () => {
		const target = mountOverlay({ type: 'morph', durationMs: 2000 } as PptxSlideTransition, {
			from: makeSlide('6', [
				bandPicture('bg6', 'backdrop', 0, 720),
				letterColumn('col6', -916, 'Calibri Light'),
				bandPicture('strip6', 'strip', -1, 720),
				bandPicture('band6', 'band', -1, 196),
			]),
			to: makeSlide('7', [
				bandPicture('bg7', 'backdrop', 0, 720),
				bandPicture('strip7', 'strip', -3, 720),
				letterColumn('col7', -158, 'Calibri'),
				bandPicture('band7', 'band', -2, 196),
			]),
		});

		const zOf = (node: Element | null | undefined): number =>
			Number(/z-index:\s*(\d+)/u.exec(node?.getAttribute('style') ?? '')?.[1]);
		const layerOf = (id: string): Element | null | undefined =>
			target
				.querySelector(`[data-pptx-morph-outgoing] [data-element-id="${id}"]`)
				?.closest('[data-pptx-morph-outgoing]');
		const group = target.querySelector('[data-pptx-morph-crossfade="col7"]');
		expect(group, 'the column must dissolve as its own group').not.toBeNull();
		expect(zOf(layerOf('band6'))).toBeGreaterThan(zOf(group));
		expect(zOf(layerOf('strip6'))).toBeLessThan(zOf(group));
	});
});

/** The `@keyframes` block `name` names in `css`, or undefined. */
function keyframesBlock(css: string, name: string): string | undefined {
	const start = css.indexOf(`@keyframes ${name} {`);
	if (start < 0) {
		return undefined;
	}
	const end = css.indexOf('@keyframes', start + 1);
	return css.slice(start, end < 0 ? undefined : end);
}

/** The opacity a keyframes block reaches at `progress` (0..1), linear between stops. */
function opacityAt(block: string, progress: number): number {
	const stops: { at: number; opacity: number }[] = [];
	for (const match of block.matchAll(/([\w%.,\s]+)\{([^{}]*)\}/gu)) {
		const opacity = /opacity:\s*([\d.]+)/u.exec(match[2] ?? '')?.[1];
		if (opacity === undefined) {
			continue;
		}
		for (const raw of (match[1] ?? '').split(',')) {
			const key = raw.trim();
			const at = key === 'from' ? 0 : key === 'to' ? 1 : Number.parseFloat(key) / 100;
			if (Number.isFinite(at)) {
				stops.push({ at, opacity: Number(opacity) });
			}
		}
	}
	stops.sort((a, b) => a.at - b.at);
	if (stops.length === 0) {
		return 1;
	}
	const after = stops.findIndex((stop) => stop.at >= progress);
	if (after <= 0) {
		return stops[after < 0 ? stops.length - 1 : 0]!.opacity;
	}
	const a = stops[after - 1]!;
	const b = stops[after]!;
	return a.opacity + ((b.opacity - a.opacity) * (progress - a.at)) / (b.at - a.at || 1);
}

const MORPH_SCOPES = [
	'data-pptx-morph-incoming',
	'data-pptx-morph-outgoing',
	'data-pptx-morph-lifted',
] as const;

/**
 * Every stylesheet the overlay rendered, joined. The morph rules are one of
 * possibly several `<style>` elements (a painted element may carry its own),
 * so reading only the first can miss the rule that hides or fades a copy.
 */
function overlayCss(target: HTMLElement): string {
	return [...target.querySelectorAll('style')].map((node) => node.textContent ?? '').join('\n');
}

/**
 * The opacity at `progress` of a painted element: the one the overlay's scoped
 * stylesheet gives the element, times the animation on any wrapper up to the
 * overlay (a crossfade half carries its dissolve inline).
 */
function paintedOpacity(target: HTMLElement, node: HTMLElement, progress: number): number {
	const css = overlayCss(target);
	const id = node.dataset.elementId ?? '';
	let opacity = 1;
	for (const scope of MORPH_SCOPES) {
		if (!node.closest(`[${scope}]`)) {
			continue;
		}
		const rule = `[${scope}] [data-element-id="${id}"] { animation: `;
		const start = css.indexOf(rule);
		if (start >= 0) {
			const name = css.slice(start + rule.length).split(/\s/u)[0] ?? '';
			const block = keyframesBlock(css, name);
			if (block) {
				opacity *= opacityAt(block, progress);
			}
		}
	}
	for (let wrapper = node.parentElement; wrapper && wrapper !== target;) {
		const name = /animation:\s*([\w-]+)/u.exec(wrapper.getAttribute('style') ?? '')?.[1];
		const block = name ? keyframesBlock(css, name) : undefined;
		if (block) {
			opacity *= opacityAt(block, progress);
		}
		wrapper = wrapper.parentElement;
	}
	return opacity;
}

/** A 17 -> 18-shaped paragraph box: three paragraphs, the second one resized. */
function parkedParagraphs(id: string, x: number, secondSize?: number): unknown {
	return {
		id,
		type: 'text',
		name: 'TextBox 9',
		x,
		y: 120,
		width: 300,
		height: 240,
		text: 'Parked opener\nSecond paragraph\nClosing line',
		textSegments: [
			{ text: 'Parked opener', style: { fontFamily: 'Arial', fontSize: 18 } },
			{ text: '\n' },
			{
				text: 'Second paragraph',
				style: {
					fontFamily: 'Arial',
					...(secondSize === undefined ? {} : { fontSize: secondSize }),
				},
			},
			{ text: '\n' },
			{ text: 'Closing line', style: { fontFamily: 'Arial', fontSize: 18 } },
		],
	};
}

const panel = (id: string): unknown => ({
	id,
	type: 'shape',
	name: 'Panel',
	x: -600,
	y: 0,
	width: 1200,
	height: 720,
	shapeStyle: { fillMode: 'solid', fillColor: '#123456' },
});

// Regression (issue #141, slides 17 -> 18 and back): a paragraph box parked off
// the left edge whose second paragraph changes size glides onto the slide.
// Cross-dissolving its two layouts painted the paragraph twice, a ghost of one
// layout over the other, for the whole flight.
describe('off-stage paragraph flight paints its text once', () => {
	it.each([
		[
			'forward',
			makeSlide('17', [panel('p17'), parkedParagraphs('t17', -400, 14)]),
			makeSlide('18', [parkedParagraphs('t18', 875)]),
		],
		[
			'backward',
			makeSlide('18', [panel('p18'), parkedParagraphs('t18', 875)]),
			makeSlide('17', [parkedParagraphs('t17', -400, 14)]),
		],
	])('%s', (_direction, from, to) => {
		const target = mountOverlay({ type: 'morph', durationMs: 1000 } as PptxSlideTransition, {
			from,
			to,
		});
		const copies = [...target.querySelectorAll<HTMLElement>('[data-element-id]')].filter(
			(node) =>
				/^t1[78]$/u.test(node.dataset.elementId ?? '') &&
				Boolean(node.textContent?.includes('Second paragraph')),
		);
		expect(copies.length, 'the stage copy and the ghost are both painted').toBeGreaterThan(1);
		const visible = copies.filter((node) => paintedOpacity(target, node, 0.5) > 0.01);
		expect(visible).toHaveLength(1);
	});
});

// Regression (issue #141, slides 8 -> 9): a picture faded in through its
// transparency tweens its `<img>`. The shared rule is a DESCENDANT selector
// (`[data-element-id] img`), so it only plays while the picture renders its
// `<img>` inside the node carrying the element id.
describe('picture transparency morph', () => {
	it('renders the `<img>` the alpha keyframes target', () => {
		const faded = (id: string, alphaModFix: number): unknown => ({
			...(bandPicture(id, 'faded', 0, 720) as object),
			imageEffects: { alphaModFix },
		});
		const target = mountOverlay({ type: 'morph', durationMs: 2000 } as PptxSlideTransition, {
			from: makeSlide('8', [bandPicture('back8', 'backdrop', 0, 720), faded('faded8', 0)]),
			to: makeSlide('9', [bandPicture('back9', 'backdrop', 0, 720), faded('faded9', 100)]),
		});
		const css = overlayCss(target);
		const selectors = [...css.matchAll(/^(\[[^{\n]*\] img) \{ animation: /gmu)].map(
			(match) => match[1]!,
		);
		// The stage copy tweens in through the incoming layer's rule.
		expect(selectors).toContain('[data-pptx-morph-incoming] [data-element-id="faded9"] img');
		expect(
			target.querySelector('[data-pptx-morph-incoming] [data-element-id="faded9"] img'),
		).toBeInstanceOf(HTMLImageElement);
		// Every rule whose element is painted in that scope reaches its `<img>`.
		for (const selector of selectors) {
			if (target.querySelector(selector.replace(/ img$/u, ''))) {
				expect(target.querySelector(selector), selector).toBeInstanceOf(HTMLImageElement);
			}
		}
	});
});
