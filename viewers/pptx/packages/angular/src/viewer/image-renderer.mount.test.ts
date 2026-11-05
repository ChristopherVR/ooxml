/**
 * `ImageRendererComponent` mounted for real (TestBed), for the one structural
 * fact a morph needs from it: the `<img>` sits INSIDE the node carrying
 * `data-element-id`.
 *
 * A picture's crop and transparency morphs are injected as descendant rules,
 * `[data-element-id="id"] img { animation: ... }` (issue #141's slides 8 to 9
 * fade a picture in through its `alphaModFix`). Were the marker moved onto the
 * `<img>`, or the picture painted without one, the rule would silently match
 * nothing and the picture would pop in when the transition ends.
 *
 * Plain JIT has no signal-input transform, so the inputs are written straight
 * into their signals (see `presentation-transition-overlay.mount.test.ts`).
 */
import type { InputSignal } from '@angular/core';
import { SIGNAL } from '@angular/core/primitives/signals';
import { TestBed } from '@angular/core/testing';
import type { ComponentFixture } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { buildMorphScopedCss, buildMorphTransitionPlan } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { ImageRendererComponent } from './image-renderer.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
});

let fixture: ComponentFixture<ImageRendererComponent> | undefined;

afterEach(() => {
	fixture?.destroy();
	fixture = undefined;
	TestBed.resetTestingModule();
	document.body.innerHTML = '';
});

const PIXEL = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
const FADED = 'ppt/media/faded.png';

function picture(id: string, alphaModFix: number): PptxElement {
	return {
		id,
		name: 'Picture 4',
		type: 'picture',
		x: 0,
		y: 0,
		width: 1280,
		height: 720,
		imagePath: FADED,
		imageEffects: { alphaModFix },
	} as PptxElement;
}

function slide(id: string, elements: PptxElement[]): PptxSlide {
	return { id, slideNumber: 1, elements } as unknown as PptxSlide;
}

/**
 * Write a value straight into one of the component's signal inputs. Written
 * against `@angular/core` 22.2.4, whose `SIGNAL` node exposes
 * `applyValueToInputSignal`; `setInput` cannot reach signal inputs under JIT.
 */
function setSignalInput(component: ImageRendererComponent, name: string, value: unknown): void {
	const node = (component[name as keyof ImageRendererComponent] as InputSignal<unknown>)[
		SIGNAL
	] as { applyValueToInputSignal: (node: unknown, value: unknown) => void };
	node.applyValueToInputSignal(node, value);
}

function render(element: PptxElement): HTMLElement {
	TestBed.configureTestingModule({ imports: [ImageRendererComponent] });
	fixture = TestBed.createComponent(ImageRendererComponent);
	setSignalInput(fixture.componentInstance, 'element', element);
	setSignalInput(fixture.componentInstance, 'mediaDataUrls', new Map([[FADED, PIXEL]]));
	const root = fixture.nativeElement as HTMLElement;
	document.body.appendChild(root);
	fixture.detectChanges();
	return root;
}

describe('imageRendererComponent under a picture transparency morph (issue #141)', () => {
	const from = picture('s8-faded', 0);
	const to = picture('s9-faded', 100);
	const plan = buildMorphTransitionPlan(slide('8', [from]), slide('9', [to]), 1000)!;

	it('paints the <img> the alpha keyframe rule selects', () => {
		const css = buildMorphScopedCss(plan, '', 'incoming');
		const selector = /(\[data-element-id="s9-faded"\] img) \{ animation: /u.exec(css)?.[1];
		expect(selector, 'the plan tweens the arriving picture through its <img>').toBeDefined();

		const root = render(to);
		const img = root.querySelector(selector!);
		expect(img, 'the rule must reach the rendered <img>').not.toBeNull();
		expect(img!.tagName).toBe('IMG');
		expect(img!.getAttribute('src')).toBe(PIXEL);
		// The marker stays on the container, not on the img itself.
		expect(img!.hasAttribute('data-element-id')).toBeFalsy();
	});
});
