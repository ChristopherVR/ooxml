// @vitest-environment happy-dom
/**
 * Issue #141 regressions, rendered the way a running React show paints them:
 * `usePresentationMode` drives the navigation, `PresentationStage` paints the
 * arriving slide with the stylesheet and element states the hook hands it,
 * and `PresentationTransitionOverlay` is mounted over it under the same key
 * `ViewerCanvasArea` gives it (`<outgoing>-<incoming>`). `ViewerCanvasArea`
 * itself needs the whole viewer state, so this harness mirrors only its
 * overlay mount.
 */
import type { PptxElement, PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import React, { act, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { usePresentationMode } from '../hooks/usePresentationMode';
import type { UsePresentationModeResult } from '../hooks/usePresentationMode';
import { PresentationStage } from './presentation/PresentationStage';
import { PresentationTransitionOverlay } from './PresentationTransitionOverlay';

const PIXEL = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
const canvasSize = { width: 1280, height: 720 };

let container: HTMLDivElement;
let root: Root;
let latest: UsePresentationModeResult | undefined;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	vi.useFakeTimers();
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	latest = undefined;
	vi.useRealTimers();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

interface HarnessProps {
	slides: PptxSlide[];
	/** Drop the overlay's key, to show what the key is there for. */
	unkeyed?: boolean;
}

function Harness({ slides, unkeyed = false }: HarnessProps): React.ReactElement {
	const containerRef = useRef<HTMLDivElement | null>(null);
	const presentation = usePresentationMode({
		mode: 'present',
		slides,
		templateElementsBySlideId: {},
		visibleSlideIndexes: slides.map((_, index) => index),
		activeSlideIndex: 0,
		containerRef,
		onSetMode: () => {},
		onSetActiveSlideIndex: () => {},
	});
	latest = presentation;
	const overlay = presentation.transitionOverlay;
	const outgoing = overlay ? slides[overlay.outgoingSlideIndex] : undefined;
	return (
		<div ref={containerRef}>
			<PresentationStage
				activeSlide={slides[presentation.presentationSlideIndex]}
				templateElements={[]}
				canvasSize={canvasSize}
				mediaDataUrls={new Map()}
				presentationElementStates={presentation.presentationElementStates}
				presentationKeyframesCss={presentation.presentationKeyframesCss}
			>
				{(scale) =>
					overlay && outgoing ? (
						<PresentationTransitionOverlay
							key={
								unkeyed ? 'overlay' : `${overlay.outgoingSlideIndex}-${overlay.incomingSlideIndex}`
							}
							outgoingSlide={outgoing}
							templateElements={[]}
							canvasSize={canvasSize}
							transition={overlay.transition}
							durationMs={overlay.durationMs}
							scale={scale}
							morphPlan={presentation.morphPlan}
							incomingSlide={slides[overlay.incomingSlideIndex]}
							onComplete={presentation.handleTransitionOverlayComplete}
						/>
					) : null
				}
			</PresentationStage>
		</div>
	);
}

function show(slides: PptxSlide[], unkeyed = false): void {
	act(() => {
		root.render(<Harness slides={slides} unkeyed={unkeyed} />);
	});
}

function advance(): void {
	act(() => {
		latest?.movePresentationSlide(1);
	});
}

function tick(ms: number): void {
	act(() => {
		vi.advanceTimersByTime(ms);
	});
}

// ---------------------------------------------------------------------------
// Reading what is painted at a given progress
// ---------------------------------------------------------------------------

/** Every `@keyframes` block in the painted stylesheets, by name. */
function keyframeBlocks(css: string): Map<string, string> {
	const blocks = new Map<string, string>();
	const head = /@keyframes\s+([\w-]+)\s*\{/gu;
	for (let match = head.exec(css); match; match = head.exec(css)) {
		let depth = 1;
		let index = head.lastIndex;
		while (depth > 0 && index < css.length) {
			if (css[index] === '{') {
				depth += 1;
			}
			if (css[index] === '}') {
				depth -= 1;
			}
			index += 1;
		}
		blocks.set(match[1]!, css.slice(head.lastIndex, index - 1));
	}
	return blocks;
}

/** The opacity a keyframes block gives at `progress` (0..1), linear between stops. */
function opacityAt(block: string, progress: number): number {
	const stops: { at: number; opacity: number }[] = [];
	const stop = /([\d.%,\s]+|from|to)\s*\{([^}]*)\}/gu;
	for (let match = stop.exec(block); match; match = stop.exec(block)) {
		const opacity = /(?:^|[;\s])opacity:\s*([\d.]+)/u.exec(match[2]!);
		if (!opacity) {
			continue;
		}
		for (const raw of match[1]!.split(',')) {
			const key = raw.trim();
			const at = key === 'from' ? 0 : key === 'to' ? 1 : Number.parseFloat(key) / 100;
			if (Number.isFinite(at)) {
				stops.push({ at, opacity: Number(opacity[1]) });
			}
		}
	}
	if (stops.length === 0) {
		return 1;
	}
	stops.sort((a, b) => a.at - b.at);
	const before = [...stops].reverse().find((entry) => entry.at <= progress) ?? stops[0]!;
	const after = stops.find((entry) => entry.at >= progress) ?? stops.at(-1)!;
	if (after.at === before.at) {
		return before.opacity;
	}
	const t = (progress - before.at) / (after.at - before.at);
	return before.opacity + (after.opacity - before.opacity) * t;
}

/** The animation names applied to `node`, inline or through a painted rule. */
function animationNames(node: HTMLElement, css: string): string[] {
	const shorthands = [node.style.animation];
	const rule = /([^{}]+)\{\s*animation:\s*([^;}]+);?\s*\}/gu;
	for (let match = rule.exec(css); match; match = rule.exec(css)) {
		const selector = match[1]!.trim();
		if (selector.startsWith('@') || /^(from|to|[\d.]+%)/u.test(selector)) {
			continue;
		}
		try {
			if (node.matches(selector)) {
				shorthands.push(match[2]!);
			}
		} catch {
			// Not a selector this DOM can evaluate (a keyframes stop, say).
		}
	}
	return shorthands
		.flatMap((shorthand) => (shorthand ?? '').split(','))
		.map((shorthand) => shorthand.trim().split(/\s+/u)[0] ?? '')
		.filter((name) => name !== '' && name !== 'none');
}

/** Whether `node` (or anything above it) is painted at all at `progress`. */
function visibleAt(node: HTMLElement, css: string, progress: number): boolean {
	const blocks = keyframeBlocks(css);
	for (let at: HTMLElement | null = node; at && at !== container; at = at.parentElement) {
		if (at.style.display === 'none' || at.style.visibility === 'hidden') {
			return false;
		}
		if (at.style.opacity !== '' && Number(at.style.opacity) === 0) {
			return false;
		}
		for (const name of animationNames(at, css)) {
			const block = blocks.get(name);
			if (block !== undefined && opacityAt(block, progress) <= 0) {
				return false;
			}
		}
	}
	return true;
}

/** The painted copies carrying `text`: the outermost element containers holding it. */
function copiesOf(text: string): HTMLElement[] {
	// The stage marks an element with `data-element-id`; the overlay's static
	// renderer with `data-static-element-type` (inside its `data-pptx-morph-*` wrapper).
	const element = '[data-element-id], [data-static-element-type]';
	return [...container.querySelectorAll<HTMLElement>(element)].filter(
		(node) =>
			node.textContent?.includes(text) &&
			!node.parentElement?.closest(element)?.textContent?.includes(text),
	);
}

function paintedCss(): string {
	return [...container.querySelectorAll('style')].map((style) => style.textContent).join('\n');
}

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

/** A full-height panel that only the leaving slide has, so the box leaves a ghost. */
function panel(id: string): PptxElement {
	return {
		id,
		type: 'shape',
		name: 'Panel',
		x: -600,
		y: 0,
		width: 1200,
		height: 720,
		shapeStyle: { fillMode: 'solid', fillColor: '#123456' },
	} as unknown as PptxElement;
}

/**
 * The parked three-paragraph box of issue #141's slides 17 -> 18: off the left
 * edge on 17 with its second paragraph at 14 pt, at 656 pt on 18 with that
 * paragraph at the inherited size.
 */
function parkedBox(id: string, x: number, secondSize?: number): PptxElement {
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
	} as unknown as PptxElement;
}

const morph: PptxSlideTransition = { type: 'morph', durationMs: 1000 };

function deck(id: string, elements: PptxElement[], transition?: PptxSlideTransition): PptxSlide {
	return { id, rId: `r-${id}`, elements, ...(transition ? { transition } : {}) } as PptxSlide;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('issue #141: a parked paragraph changing size (slides 17 -> 18)', () => {
	const cases = [
		{
			direction: 'forward',
			slides: [
				deck('s17', [panel('p17'), parkedBox('t17', -400, 14)]),
				deck('s18', [parkedBox('t18', 875)], morph),
			],
		},
		{
			direction: 'backward',
			slides: [
				deck('s18', [panel('p18'), parkedBox('t18', 875)]),
				deck('s17', [parkedBox('t17', -400, 14)], morph),
			],
		},
	];

	it.each(cases)('paints the paragraph once at mid-morph ($direction)', ({ slides }) => {
		show(slides);
		advance();
		expect(latest?.morphPlan, 'the morph must be planned').toBeDefined();

		const copies = copiesOf('Second paragraph');
		const css = paintedCss();
		// The stage's arriving copy and the overlay's ghost are both mounted...
		expect(copies.length).toBeGreaterThanOrEqual(2);
		// ...but only one of them is visible halfway through.
		expect(copies.filter((node) => visibleAt(node, css, 0.5))).toHaveLength(1);
	});
});

describe('issue #141: a picture fading in through its transparency (slides 8 -> 9)', () => {
	function picture(id: string, alphaModFix?: number, extra: Partial<PptxElement> = {}) {
		return {
			id,
			name: id,
			type: 'picture',
			x: 0,
			y: 0,
			width: 1280,
			height: 720,
			imagePath: 'ppt/media/faded.png',
			imageData: PIXEL,
			...(alphaModFix === undefined ? {} : { imageEffects: { alphaModFix } }),
			...extra,
		} as unknown as PptxElement;
	}

	it('renders each copy as an <img> inside its element container, where the alpha rule reaches it', () => {
		show([
			deck('s8', [
				picture('s8-band', undefined, { y: 456, height: 264, imagePath: 'ppt/media/band.png' }),
				picture('s8-faded', 0),
			]),
			deck('s9', [picture('s9-faded', 100)], morph),
		]);
		advance();

		const plan = latest!.morphPlan!;
		const stageTrack = plan.incomingImageAnimations.get('s9-faded');
		const ghostTrack = plan.outgoingImageAnimations.get('s8-faded');
		expect(stageTrack).toBeDefined();
		expect(ghostTrack).toBeDefined();

		// The stage copy: the shared descendant rule has to find an <img>.
		const css = paintedCss();
		expect(css).toContain(`[data-element-id="s9-faded"] img { animation: ${stageTrack}; }`);
		const stageImg = container.querySelector<HTMLElement>(
			'[data-pptx-presentation-stage] [data-element-id="s9-faded"] img',
		);
		expect(stageImg, 'the stage picture must be an <img> inside its container').not.toBeNull();
		expect(animationNames(stageImg!, css)).toContain(stageTrack!.split(' ')[0]);

		// The ghost: the same channel, carried inline on its own <img> (the
		// overlay's static renderer marks the element with
		// `data-static-element-type`, not `data-element-id`).
		const ghostImg = container.querySelector<HTMLElement>(
			'[data-pptx-morph-outgoing="s8-faded"] [data-static-element-type="picture"] img',
		);
		expect(ghostImg, 'the ghost picture must be an <img> inside its container').not.toBeNull();
		expect(animationNames(ghostImg!, css)).toContain(ghostTrack!.split(' ')[0]);
	});
});

describe('consecutive transitions', () => {
	// A fade still running when the presenter moves on must not take the next
	// transition's overlay down with its own completion timer: the morph would
	// then cut to its end state partway through.
	it('plays the next transition in full when it starts during the previous one', () => {
		const fade: PptxSlideTransition = { type: 'fade', durationMs: 1000 };
		show([
			deck('a', [panel('a-panel')]),
			deck('b', [parkedBox('b-box', 100)], fade),
			deck('c', [parkedBox('c-box', 500)], morph),
		]);
		advance();
		expect(latest?.transitionOverlay?.transition.type).toBe('fade');

		tick(600);
		advance();
		expect(latest?.transitionOverlay?.transition.type).toBe('morph');

		// Past the point the fade's own timer would have fired (1000 + 50 ms).
		tick(900);
		expect(latest?.transitionOverlay?.transition.type).toBe('morph');
		expect(container.querySelector('[data-pptx-transition-morph]')).not.toBeNull();

		// And the morph's own completion still lands.
		tick(200);
		expect(latest?.transitionOverlay).toBeNull();
		expect(container.querySelector('[data-pptx-transition-overlay]')).toBeNull();
	});
});
