/**
 * PresentationTransitionOverlayComponent mounted for real (TestBed), for what
 * only a running component shows: when `complete` fires, and what its layers
 * plus the morph rules it injects at document level actually paint.
 *
 * Plain JIT here has no signal-input transform, so child components bound in
 * a template never receive their signal inputs (`pptx-slide-canvas` throws
 * NG0950). The component's own logic (its inputs, derived state, effects,
 * completion timer and injected stylesheet) runs unchanged; only its template
 * is swapped for a stand-in that paints each layer the way
 * `MorphExtraLayersComponent` does (same scope attributes, same `[ngStyle]`),
 * with every element reduced to its `data-element-id` node and its text. The
 * real template's bindings are asserted as source in
 * `presentation-transition-overlay.component.test.ts`.
 */
import { NgStyle } from '@angular/common';
import type { InputSignal } from '@angular/core';
import { SIGNAL } from '@angular/core/primitives/signals';
import { TestBed } from '@angular/core/testing';
import type { ComponentFixture } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import type { PptxElement, PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { resolveViewerComponentResources } from './component-resources.test-support';
import { PresentationTransitionOverlayComponent } from './presentation-transition-overlay.component';

beforeAll(async () => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

let fixture: ComponentFixture<PresentationTransitionOverlayComponent> | undefined;

afterEach(() => {
	// Destroying the component also removes the morph `<style>` it injected.
	fixture?.destroy();
	fixture = undefined;
	TestBed.resetTestingModule();
	vi.useRealTimers();
	document.body.innerHTML = '';
});

/** Each layer as `MorphExtraLayersComponent` paints it, elements reduced to their marked node. */
const STAND_IN_TEMPLATE = `
	@for (layer of overlayLayers(); track layer.key) {
		<div
			[attr.data-pptx-morph-outgoing]="layer.side === 'outgoing' ? 'true' : null"
			[attr.data-pptx-morph-lifted]="layer.side === 'incoming' ? 'true' : null"
			[ngStyle]="layer.style"
		>
			@for (element of layer.slide.elements; track element.id) {
				<div [attr.data-element-id]="element.id">{{ element.text }}</div>
			}
		</div>
	}
	@for (group of crossfadeGroups(); track group.key) {
		<div [attr.data-pptx-morph-crossfade]="group.key" [ngStyle]="group.style">
			<div data-pptx-morph-outgoing="true" [ngStyle]="group.outgoingStyle">
				@for (element of group.outgoing.elements; track element.id) {
					<div [attr.data-element-id]="element.id">{{ element.text }}</div>
				}
			</div>
			<div data-pptx-morph-lifted="true" [ngStyle]="group.incomingStyle">
				@for (element of group.incoming.elements; track element.id) {
					<div [attr.data-element-id]="element.id">{{ element.text }}</div>
				}
			</div>
		</div>
	}
`;

interface OverlayInputs {
	outgoingSlide: PptxSlide;
	incomingSlide?: PptxSlide;
	transition: PptxSlideTransition;
	durationMs?: number;
}

/**
 * Write a value straight into one of the component's signal inputs, which the
 * component's derived state reads. `setInput` cannot reach them under plain JIT.
 */
function setSignalInput(
	component: PresentationTransitionOverlayComponent,
	name: keyof OverlayInputs | 'canvasSize',
	value: unknown,
): void {
	const node = (component[name] as InputSignal<unknown>)[SIGNAL] as {
		applyValueToInputSignal: (node: unknown, value: unknown) => void;
	};
	node.applyValueToInputSignal(node, value);
}

function mount(inputs: OverlayInputs): {
	root: HTMLElement;
	completed: () => number;
	set: (name: keyof OverlayInputs, value: unknown) => void;
} {
	TestBed.configureTestingModule({
		imports: [PresentationTransitionOverlayComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	TestBed.overrideComponent(PresentationTransitionOverlayComponent, {
		set: {
			templateUrl: '',
			template: STAND_IN_TEMPLATE,
			styleUrl: '',
			styles: '',
			imports: [NgStyle],
		},
	});
	fixture = TestBed.createComponent(PresentationTransitionOverlayComponent);
	const component = fixture.componentInstance;
	setSignalInput(component, 'canvasSize', { width: 1280, height: 720 });
	for (const [name, value] of Object.entries(inputs)) {
		setSignalInput(component, name as keyof OverlayInputs, value);
	}
	let count = 0;
	component.complete.subscribe(() => {
		count += 1;
	});
	const root = fixture.nativeElement as HTMLElement;
	document.body.appendChild(root);
	fixture.detectChanges();
	return {
		root,
		completed: () => count,
		set: (name, value) => {
			setSignalInput(component, name, value);
			fixture!.detectChanges();
		},
	};
}

function slide(id: string, elements: PptxElement[]): PptxSlide {
	return { id, elements, backgroundColor: '#ffffff' } as unknown as PptxSlide;
}

function rect(id: string, x: number): PptxElement {
	return { id, type: 'shape', name: 'Rect', x, y: 10, width: 100, height: 50 } as PptxElement;
}

// Issue #141: the overlay is reused from one transition to the next, so a
// fade's completion timer must not survive into the morph that replaces it.
// Left armed it would tear the morph down after the fade's 700 ms, mid-glide.
describe('presentationTransitionOverlayComponent completion timer', () => {
	it('re-arms for a new transition instead of completing on the old timer', () => {
		vi.useFakeTimers();
		const overlay = mount({
			outgoingSlide: slide('a', [rect('a-box', 10)]),
			incomingSlide: slide('b', [rect('b-box', 400)]),
			transition: { type: 'fade', durationMs: 700 },
		});

		vi.advanceTimersByTime(500);
		overlay.set('transition', { type: 'morph', durationMs: 2000 });

		// The fade's own timer would have fired at 750 ms.
		vi.advanceTimersByTime(1000);
		expect(overlay.completed()).toBe(0);

		vi.advanceTimersByTime(1100);
		expect(overlay.completed()).toBe(1);
		vi.advanceTimersByTime(5000);
		expect(overlay.completed()).toBe(1);
	});
});

// ---------------------------------------------------------------------------
// What is visible at a given moment of the morph
// ---------------------------------------------------------------------------

/** Opacity a `@keyframes` block holds at `progress` (0..1), interpolated linearly. */
function keyframeOpacityAt(css: string, name: string, progress: number): number {
	const block = css.split(`@keyframes ${name} {`)[1]?.split('@keyframes')[0] ?? '';
	const stops: { at: number; opacity: number }[] = [];
	for (const match of block.matchAll(/([\d.%,\sfromto]+)\{([^}]*)\}/gu)) {
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
		return stops[Math.max(after, 0)]!.opacity;
	}
	const a = stops[after - 1]!;
	const b = stops[after]!;
	return a.opacity + ((b.opacity - a.opacity) * (progress - a.at)) / (b.at - a.at || 1);
}

/**
 * The animation the document's stylesheets give `node`: the rule with the most
 * attribute selectors wins (a scoped rule outranks the unscoped stage rule),
 * then the last one; an inline `animation` beats both.
 */
function animationNameOf(node: HTMLElement, css: string): string | undefined {
	const inline =
		node.style.animation || node.getAttribute('style')?.match(/animation:\s*([^;]+)/u)?.[1];
	if (inline) {
		return inline.trim().split(/\s+/u)[0];
	}
	let best: { weight: number; name: string } | undefined;
	for (const match of css.matchAll(/([^{}@]+?)\{\s*animation:\s*([^\s;]+)[^}]*\}/gu)) {
		const selector = (match[1] ?? '').trim();
		if (!selector || !node.matches(selector)) {
			continue;
		}
		const weight = (selector.match(/\[/gu) ?? []).length;
		if (!best || weight >= best.weight) {
			best = { weight, name: match[2]! };
		}
	}
	return best?.name;
}

/** Whether `node` or an ancestor is painted at `progress` of the morph. */
function visibleAt(node: HTMLElement, css: string, progress: number): boolean {
	for (let at: HTMLElement | null = node; at; at = at.parentElement) {
		const name = animationNameOf(at, css);
		if (name && keyframeOpacityAt(css, name, progress) <= 0.001) {
			return false;
		}
	}
	return true;
}

/** Every `data-element-id` node carrying `text` that is painted at `progress`. */
function visibleCopies(text: string, progress: number): HTMLElement[] {
	const css = [...document.querySelectorAll('style')].map((style) => style.textContent).join('\n');
	return [...document.querySelectorAll<HTMLElement>('[data-element-id]')].filter(
		(node) => node.textContent?.includes(text) && visibleAt(node, css, progress),
	);
}

// Issue #141, slides 17 -> 18 and back: a three-paragraph text box parked off
// the left edge glides onto the slide while its second paragraph changes from
// 14 pt to the inherited size. Cross-dissolving the two layouts painted the
// paragraph twice, a visible ghost beside the arriving text.
describe('presentationTransitionOverlayComponent off-stage size change', () => {
	const BODY = 'Parked opener\nSecond paragraph\nClosing line';
	function textBox(id: string, x: number, secondSize?: number): PptxElement {
		return {
			id,
			type: 'text',
			name: 'TextBox 9',
			x,
			y: 120,
			width: 300,
			height: 240,
			text: BODY,
			textSegments: [
				{ text: 'Parked opener', style: { fontFamily: 'Arial', fontSize: 18 } },
				{ text: '\n' },
				{
					text: 'Second paragraph',
					style: { fontFamily: 'Arial', ...(secondSize ? { fontSize: secondSize } : {}) },
				},
				{ text: '\n' },
				{ text: 'Closing line', style: { fontFamily: 'Arial', fontSize: 18 } },
			],
		} as unknown as PptxElement;
	}
	const panel = (id: string): PptxElement =>
		({
			id,
			type: 'shape',
			name: 'Panel',
			x: -600,
			y: 0,
			width: 1200,
			height: 720,
			shapeStyle: { fillMode: 'solid', fillColor: '#123456' },
		}) as unknown as PptxElement;

	/** The live stage, a sibling of the overlay that the unscoped rules reach. */
	function paintStage(element: PptxElement): void {
		const stage = document.createElement('div');
		const node = document.createElement('div');
		node.dataset.elementId = element.id;
		node.textContent = (element as { text?: string }).text ?? '';
		stage.appendChild(node);
		document.body.appendChild(stage);
	}

	it.each([
		{
			direction: 'forward',
			from: slide('s17', [panel('p17'), textBox('t17', -400, 14)]),
			to: slide('s18', [textBox('t18', 875)]),
		},
		{
			direction: 'backward',
			from: slide('s18', [panel('p18'), textBox('t18', 875)]),
			to: slide('s17', [textBox('t17', -400, 14)]),
		},
	])('paints the text once mid-morph ($direction)', ({ from, to }) => {
		paintStage(to.elements.at(-1)!);
		mount({
			outgoingSlide: from,
			incomingSlide: to,
			transition: { type: 'morph', durationMs: 1000 },
		});

		// Both copies exist (the ghost in the overlay, the arrival on the stage)...
		const copies = [...document.querySelectorAll<HTMLElement>('[data-element-id]')].filter((node) =>
			node.textContent?.includes('Second paragraph'),
		);
		expect(copies).toHaveLength(2);
		// ...but exactly one of them is painted halfway through.
		expect(visibleCopies('Second paragraph', 0.5)).toHaveLength(1);
	});
});
