/**
 * `*pptxElementBoundary` (the element host's try/catch): a failing element is omitted alone and
 * reported through the `ElementRenderErrors` scope; the same failed object is not retried, a
 * replacement object or a new id is; a throwing reporter falls back to `console.error`.
 */
import { Component, inject, signal } from '@angular/core';
import type { ComponentFixture } from '@angular/core/testing';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import { ELEMENT_RENDER_FAULT_KEY } from 'ooxml-ui/pptx';
import type { ElementRenderErrorHandler } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { resolveViewerComponentResources } from './component-resources.test-support';
import { ElementRenderErrors } from './element-boundary.directive';
import { RibbonLayoutArtworkComponent } from './ribbon-layout-artwork';
import { declareSignalFields } from './signal-inputs.test-support';

const FAULT = Symbol.for(ELEMENT_RENDER_FAULT_KEY);
const setFault = (id: string | undefined) => {
	(globalThis as Record<symbol, unknown>)[FAULT] = id;
};
const shape = (id: string, x = 0): PptxElement =>
	({
		id,
		type: 'shape',
		shapeType: 'rect',
		x,
		y: 0,
		width: 100,
		height: 50,
		text: `<${id}>`,
	}) as PptxElement;

const reporter: { current: ElementRenderErrorHandler | undefined } = { current: undefined };

@Component({
	selector: 'test-boundary-host',
	standalone: true,
	imports: [RibbonLayoutArtworkComponent],
	providers: [ElementRenderErrors],
	template: `<pptx-ribbon-layout-artwork [elements]="elements()" />`,
})
class BoundaryHostComponent {
	readonly elements = signal<PptxElement[]>([]);
	constructor() {
		inject(ElementRenderErrors, { self: true }).reporter = () => reporter.current;
	}
}

beforeAll(async () => {
	declareSignalFields(BoundaryHostComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

let fixture: ComponentFixture<BoundaryHostComponent> | undefined;
afterEach(() => {
	fixture?.destroy();
	fixture = undefined;
	TestBed.resetTestingModule();
	reporter.current = undefined;
	Reflect.deleteProperty(globalThis, FAULT);
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

function create(elements: PptxElement[]) {
	TestBed.configureTestingModule({ providers: [provideTranslateService({ fallbackLang: 'en' })] });
	fixture = TestBed.createComponent(BoundaryHostComponent);
	document.body.append(fixture.nativeElement);
	fixture.componentInstance.elements.set(elements);
	fixture.detectChanges();
	// The layout artwork drops element ids (a static preview), so read each element's text.
	const ids = () =>
		[...((fixture!.nativeElement as HTMLElement).textContent ?? '').matchAll(/<(\w+)>/gu)]
			.map((match) => match[1])
			.filter((id, index, all) => all.indexOf(id) === index);
	const set = (next: PptxElement[]) => {
		fixture!.componentInstance.elements.set(next);
		fixture!.detectChanges();
	};
	return { ids, set };
}

describe('Angular element boundary', () => {
	it('omits only the failing element, then retries a replacement object or a new id', () => {
		setFault('bad');
		const report = vi.fn<ElementRenderErrorHandler>();
		reporter.current = report;
		const bad = shape('bad', 120);
		const { ids, set } = create([shape('good'), bad]);
		expect(ids()).toEqual(['good']);
		expect(report).toHaveBeenCalledTimes(1);
		expect(report).toHaveBeenCalledWith(expect.any(Error), {
			elementId: 'bad',
			elementType: 'shape',
		});
		// The same failed object is not retried when the list re-renders.
		set([shape('good'), bad]);
		expect(report).toHaveBeenCalledTimes(1);
		// A replacement object is retried (and fails again while the fault holds).
		set([shape('good'), { ...bad }]);
		expect(report).toHaveBeenCalledTimes(2);
		// A healthy replacement renders.
		setFault(undefined);
		set([shape('good'), { ...bad }]);
		expect(ids()).toEqual(['good', 'bad']);
		// A changed id gets a boundary of its own.
		setFault('renamed');
		set([shape('good'), shape('renamed', 120)]);
		expect(ids()).toEqual(['good']);
		expect(report).toHaveBeenLastCalledWith(
			expect.any(Error),
			expect.objectContaining({ elementId: 'renamed' }),
		);
	});

	it('falls back to console.error when the reporter throws or is absent', () => {
		setFault('bad');
		const log = vi.spyOn(console, 'error').mockImplementation(() => {});
		reporter.current = vi.fn(() => {
			throw new Error('reporter');
		});
		const { ids } = create([shape('good'), shape('bad', 120)]);
		expect(reporter.current).toHaveBeenCalledTimes(1);
		expect(ids()).toEqual(['good']);
		expect(log).toHaveBeenCalledWith(
			expect.any(Error),
			expect.objectContaining({ elementId: 'bad' }),
		);
	});
});
