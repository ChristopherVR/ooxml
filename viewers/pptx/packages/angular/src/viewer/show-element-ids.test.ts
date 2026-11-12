/**
 * One marked node per element id while a show runs, rendered.
 *
 * Angular presents on a separate full-screen stage (`pptx-presentation-overlay`)
 * or in the presenter console (`pptx-presenter-view`), and keeps the editor
 * mounted underneath, so its state survives the show. The editing canvas kept
 * drawing the active slide there, with its `data-element-id` markers, so every
 * element on that slide answered from two nodes for the whole show. The show
 * surface is the one the user acts on; the covered canvas now draws no slide
 * until the show ends.
 */
import { TestBed } from '@angular/core/testing';
import type { ComponentFixture } from '@angular/core/testing';
import { By } from '@angular/platform-browser';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import { PptxHandler } from 'pptx-viewer-core';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { resolveViewerComponentResources } from './component-resources.test-support';
import { PowerPointViewerComponent } from './power-point-viewer.component';
import { PresentationOverlayComponent } from './presentation-overlay.component';
import { RibbonComponent } from './ribbon.component';
import { declareSignalFields } from './signal-inputs.test-support';

/** Two slides; the first holds one rectangle (the loader assigns its id). */
async function deckWithShape(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	try {
		const first = data.slides[0]!;
		first.elements = [
			...first.elements,
			{
				id: 'shape',
				type: 'shape',
				shapeType: 'rect',
				x: 100,
				y: 100,
				width: 200,
				height: 100,
				shapeStyle: { fillColor: '#ff0000' },
			} as unknown as PptxElement,
		];
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

beforeAll(async () => {
	declareSignalFields(PowerPointViewerComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

afterEach(() => TestBed.resetTestingModule());

function markedIds(root: Element): string[] {
	return [...root.querySelectorAll('[data-element-id]')].map(
		(node) => node.getAttribute('data-element-id') ?? '',
	);
}

const duplicates = (ids: readonly string[]) => ids.filter((id, at) => ids.indexOf(id) !== at);

async function settle(
	fixture: ComponentFixture<PowerPointViewerComponent>,
	done: () => boolean,
): Promise<void> {
	for (let attempt = 0; attempt < 100 && !done(); attempt += 1) {
		fixture.detectChanges();
		await fixture.whenStable();
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	}
	fixture.detectChanges();
}

type Viewer = ComponentFixture<PowerPointViewerComponent>;

/** Mount the viewer on the deck and wait for the editing canvas's markers. */
async function openViewer(): Promise<{ fixture: Viewer; root: HTMLElement; editing: string[] }> {
	TestBed.configureTestingModule({ providers: [provideTranslateService()] });
	const fixture = TestBed.createComponent(PowerPointViewerComponent);
	fixture.componentRef.setInput('content', await deckWithShape());
	const root = fixture.nativeElement as HTMLElement;
	await settle(fixture, () => markedIds(root).length > 0);
	const editing = markedIds(root);
	expect(editing.length).toBeGreaterThan(0);
	return { fixture, root, editing };
}

/** The editing canvas lives in the viewer's main column. */
const editingCanvas = (root: Element) => root.querySelector('main.pptx-ng-main')!;
const showOverlay = (root: Element) => root.querySelector('pptx-presentation-overlay');
const presenterConsole = (root: Element) => root.querySelector('pptx-presenter-view');

/** The presenter console's entry points: the show toolbar, the ribbon and the command search. */
const presenterEntries: [string, (fixture: Viewer) => Promise<void>][] = [
	[
		"the show toolbar's presenter toggle",
		async (fixture) => {
			const root = fixture.nativeElement as HTMLElement;
			fixture.componentInstance.setMode('present');
			await settle(fixture, () => showOverlay(root) !== null);
			fixture.debugElement
				.query(By.directive(PresentationOverlayComponent))
				.componentInstance.presenterViewToggle.emit();
		},
	],
	[
		"the ribbon's Presenter View",
		async (fixture) => {
			fixture.debugElement.query(By.directive(RibbonComponent)).componentInstance.presenter.emit();
		},
	],
	[
		'the Slide Show search command',
		async (fixture) => {
			fixture.componentInstance['handleCommandSearch']('slideShow.presenterView');
		},
	],
];

describe('element ids while a show runs', () => {
	it('come from the show stage only, and from the editing canvas again after it', async () => {
		const { fixture, root, editing } = await openViewer();

		fixture.componentInstance.setMode('present');
		await settle(fixture, () => showOverlay(root) !== null);
		const show = showOverlay(root);
		expect(show).not.toBeNull();
		expect(markedIds(show!)).toStrictEqual(editing);
		expect(markedIds(root)).toStrictEqual(editing);

		fixture.componentInstance.setMode('edit');
		await settle(fixture, () => showOverlay(root) === null);
		expect(markedIds(root)).toStrictEqual(editing);
	});

	it.each(presenterEntries)(
		'leave the editing canvas in the presenter view opened from %s',
		async (_entry, openPresenter) => {
			const { fixture, root, editing } = await openViewer();

			await openPresenter(fixture);
			await settle(fixture, () => presenterConsole(root) !== null);
			expect(presenterConsole(root)).not.toBeNull();
			expect(markedIds(editingCanvas(root))).toStrictEqual([]);
			expect(duplicates(markedIds(root))).toStrictEqual([]);

			fixture.componentInstance['presentationMode'].exitPresenter();
			await settle(fixture, () => markedIds(editingCanvas(root)).length > 0);
			expect(presenterConsole(root)).toBeNull();
			expect(markedIds(root)).toStrictEqual(editing);
		},
	);
});
