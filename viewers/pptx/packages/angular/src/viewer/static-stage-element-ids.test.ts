/**
 * Which Angular slide stages carry `data-element-id`, rendered.
 *
 * The editing canvas stays mounted under the presenter view and the reading
 * view, and both draw the active slide, so a marked copy there put two nodes
 * with the same id in the document. `SlideCanvasComponent` now derives the
 * markers from the shared `stageKeepsElementIds` rule: only the editing canvas,
 * the live show stage and the show's transition layers keep them, whatever the
 * call site remembers to pass.
 */
import { Component, signal } from '@angular/core';
import type { ComponentFixture, WritableSignal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { resolveViewerComponentResources } from './component-resources.test-support';
import { MobilePresenterViewComponent } from './mobile-presenter-view.component';
import { MorphExtraLayersComponent } from './morph-extra-layers.component';
import { PresenterViewComponent } from './presenter-view.component';
import { ReadingViewOverlayComponent } from './reading-view-overlay.component';
import { declareSignalFields } from './signal-inputs.test-support';
import { SlideCanvasComponent } from './slide-canvas.component';

const shape = (id: string): PptxElement =>
	({
		id,
		type: 'shape',
		shapeType: 'rect',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeStyle: { fillColor: '#ff0000' },
	}) as unknown as PptxElement;

const slides: PptxSlide[] = [
	{ id: 's1', rId: 'r1', slideNumber: 1, elements: [shape('a')] } as unknown as PptxSlide,
	{ id: 's2', rId: 'r2', slideNumber: 2, elements: [shape('b')] } as unknown as PptxSlide,
];
const canvasSize = { width: 960, height: 540 };

/** One canvas with the flags a call site may set. */
@Component({
	selector: 'test-canvas-host',
	standalone: true,
	imports: [SlideCanvasComponent],
	template: `
		<pptx-slide-canvas
			[slide]="slide"
			[canvasSize]="size"
			[autoFit]="false"
			[interactive]="interactive()"
			[presenting]="presenting()"
			[preserveElementIds]="preserveElementIds()"
		/>
	`,
})
class CanvasHostComponent {
	readonly slide = slides[0]!;
	readonly size = canvasSize;
	readonly interactive = signal(true);
	readonly presenting = signal(false);
	readonly preserveElementIds = signal(false);
}

type CanvasFlags = Partial<Record<'interactive' | 'presenting' | 'preserveElementIds', boolean>>;

beforeAll(async () => {
	for (const component of [
		CanvasHostComponent,
		PresenterViewComponent,
		MobilePresenterViewComponent,
		ReadingViewOverlayComponent,
		MorphExtraLayersComponent,
	]) {
		declareSignalFields(component, import.meta.dirname);
	}
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

afterEach(() => TestBed.resetTestingModule());

function ids(fixture: ComponentFixture<unknown>): string[] {
	const root = fixture.nativeElement as HTMLElement;
	return [...root.querySelectorAll('[data-element-id]')].map(
		(node) => node.getAttribute('data-element-id') ?? '',
	);
}

function drawnElements(fixture: ComponentFixture<unknown>): number {
	return (fixture.nativeElement as HTMLElement).querySelectorAll('.pptx-ng-element').length;
}

function mount<T>(
	component: new (...args: never[]) => T,
	inputs: Record<string, unknown>,
): ComponentFixture<T> {
	TestBed.configureTestingModule({ providers: [provideTranslateService()] });
	const fixture = TestBed.createComponent(component);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, value);
	}
	fixture.detectChanges();
	return fixture;
}

function mountCanvas(flags: CanvasFlags): ComponentFixture<CanvasHostComponent> {
	TestBed.configureTestingModule({ providers: [provideTranslateService()] });
	const fixture = TestBed.createComponent(CanvasHostComponent);
	for (const [name, value] of Object.entries(flags)) {
		fixture.componentInstance[name as keyof CanvasFlags].set(value);
	}
	fixture.detectChanges();
	return fixture;
}

describe('angular static stages draw without element ids', () => {
	it('presenter view: current and next previews', () => {
		const fixture = mount(PresenterViewComponent, { slides, currentSlideIndex: 0, canvasSize });
		expect(drawnElements(fixture)).toBeGreaterThanOrEqual(2);
		expect(ids(fixture)).toStrictEqual([]);
	});

	it('mobile presenter view: current and next previews', () => {
		const fixture = mount(MobilePresenterViewComponent, {
			slides,
			currentSlideIndex: 0,
			canvasSize,
		});
		expect(drawnElements(fixture)).toBeGreaterThanOrEqual(2);
		expect(ids(fixture)).toStrictEqual([]);
	});

	it('reading view, once its viewport is measured', () => {
		const fixture = mount(ReadingViewOverlayComponent, {
			slides,
			activeSlideIndex: 0,
			canvasSize,
		});
		// The size its ResizeObserver reports in a browser. Plain JIT resolves no
		// signal queries, so the observer never attaches here.
		const viewport = (
			fixture.componentInstance as unknown as {
				viewport: WritableSignal<{ width: number; height: number }>;
			}
		).viewport;
		viewport.set({ width: 800, height: 600 });
		fixture.detectChanges();
		const root = fixture.nativeElement as HTMLElement;
		expect(root.querySelector('[data-pptx-reading-view-stage]')).not.toBeNull();
		expect(drawnElements(fixture)).toBe(1);
		expect(ids(fixture)).toStrictEqual([]);
	});

	it('a canvas a caller marks only as non-interactive', () => {
		const fixture = mountCanvas({ interactive: false });
		expect(drawnElements(fixture)).toBe(1);
		expect(ids(fixture)).toStrictEqual([]);
	});
});

describe('angular stages that keep element ids', () => {
	it('the editing canvas', () => {
		expect(ids(mountCanvas({}))).toStrictEqual(['a']);
	});

	it('the live show stage', () => {
		expect(ids(mountCanvas({ interactive: false, presenting: true }))).toStrictEqual(['a']);
	});

	it('a transition layer', () => {
		expect(ids(mountCanvas({ interactive: false, preserveElementIds: true }))).toStrictEqual(['a']);
	});

	it('the morph layers', () => {
		const fixture = mount(MorphExtraLayersComponent, {
			overlayLayers: [{ key: 'ghost', side: 'outgoing', style: {}, slide: slides[0] }],
			crossfadeGroups: [
				{
					key: 'pair',
					style: {},
					outgoingStyle: {},
					incomingStyle: {},
					outgoing: slides[0],
					incoming: slides[1],
				},
			],
			canvasSize,
			slideBoxStyle: {},
		});
		expect(ids(fixture)).toStrictEqual(['a', 'a', 'b']);
	});
});
