/**
 * Host-supplied regions (`customization.regions`, `backstage.extraPages`,
 * `capabilities`) in the Angular binding.
 *
 * The full viewer cannot be mounted here (it pulls in every child component
 * and the canvas), so each piece is driven through the component that owns
 * it: the region host, the Options dialog frame and the File tab, inside the
 * same per-viewer services the root component provides.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
	Injector,
	runInInjectionContext,
	signal,
	ɵresolveComponentResources as resolveComponentResources,
} from '@angular/core';
import type { Type } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import type { ComponentFixture } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import type {
	PowerPointViewerAPI,
	ViewerRegionContext,
	ViewerRegionRenderer,
	ViewerSettingsRegionContext,
} from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { RibbonFileSectionComponent } from './ribbon-file-section.component';
import { SettingsDialogComponent } from './settings-dialog.component';
import { ViewerCustomizationService } from './viewer-customization.service';
import { ViewerOptionsService } from './viewer-options.service';
import { ViewerRegionHostComponent } from './viewer-region-host.component';
import { ViewerRegionService } from './viewer-region.service';
import type { ViewerRegionBase } from './viewer-region.service';

beforeAll(async () => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	// The JIT build leaves `templateUrl`/`styleUrl` unresolved; read them from disk.
	await resolveComponentResources(async (url) =>
		readFileSync(resolve(import.meta.dirname, url), 'utf8'),
	);
});
afterEach(() => TestBed.resetTestingModule());

const VIEWER_PROVIDERS = [ViewerOptionsService, ViewerCustomizationService, ViewerRegionService];

/** The viewer state the root component would attach. */
function baseState(state: { selection: string[] }): () => ViewerRegionBase {
	return () => ({
		viewer: {} as PowerPointViewerAPI,
		mode: 'edit',
		canEdit: true,
		activeSlideIndex: 0,
		slideCount: 3,
		selectedElementIds: state.selection,
		locale: 'en',
	});
}

/**
 * Create `component` inside the per-viewer services. The Vitest JIT build does
 * not wire `input()`, so each input is declared and handed a signal.
 */
async function mount<T>(
	component: Type<T>,
	values: Record<string, unknown> = {},
): Promise<ComponentFixture<T>> {
	await resolveComponentResources(async (url) =>
		readFileSync(resolve(import.meta.dirname, url), 'utf8'),
	);
	TestBed.resetTestingModule();
	TestBed.configureTestingModule({
		imports: [component],
		providers: [provideTranslateService({ fallbackLang: 'en' }), ...VIEWER_PROVIDERS],
	});
	TestBed.overrideComponent(component, { add: { inputs: Object.keys(values) } });
	const fixture = TestBed.createComponent(component);
	for (const [name, value] of Object.entries(values)) {
		fixture.componentRef.setInput(name, signal(value));
	}
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.use('en');
	return fixture;
}

const context = (selection: string[]): ViewerRegionContext => ({
	region: 'inspector',
	viewer: {} as PowerPointViewerAPI,
	mode: 'edit',
	canEdit: true,
	activeSlideIndex: 0,
	slideCount: 3,
	selectedElementIds: selection,
	locale: 'en',
	close: () => undefined,
});

describe('viewer region host', () => {
	it('mounts the renderer, delivers a new context on update and destroys on unmount', async () => {
		const update = vi.fn();
		const destroy = vi.fn();
		const render = vi.fn((container: HTMLElement) => {
			container.textContent = 'host inspector';
			return { update, destroy };
		});
		const contextSignal = signal(context([]));
		const fixture = await mount(ViewerRegionHostComponent, {
			renderer: render as ViewerRegionRenderer,
			context: context([]),
		});
		// Hand the host the live signal so the test can change the context.
		fixture.componentRef.setInput('context', contextSignal);
		fixture.detectChanges();
		expect(fixture.nativeElement.textContent).toContain('host inspector');
		expect(render).toHaveBeenCalledTimes(1);

		contextSignal.set(context(['shape-1']));
		fixture.detectChanges();
		expect(update).toHaveBeenCalledTimes(1);
		expect(update.mock.calls[0]?.[0].selectedElementIds).toStrictEqual(['shape-1']);

		fixture.destroy();
		expect(destroy).toHaveBeenCalledTimes(1);
	});
});

describe('region service', () => {
	it('returns the renderer a host registered and builds the base context', async () => {
		await mount(ViewerRegionHostComponent, { renderer: 'x', context: context([]) });
		const regions = TestBed.inject(ViewerRegionService);
		const customization = TestBed.inject(ViewerCustomizationService);
		regions.attach(baseState({ selection: ['a'] }));
		expect(regions.renderer('inspector')).toBeUndefined();
		customization.api.setRegion('inspector', 'x-inspector');
		expect(regions.renderer('inspector')).toBe('x-inspector');
		const close = vi.fn();
		const ctx = regions.context('inspector', close);
		expect(ctx).toMatchObject({ region: 'inspector', selectedElementIds: ['a'], slideCount: 3 });
		ctx?.close();
		expect(close).toHaveBeenCalledTimes(1);
	});
});

describe('unbuildable context (R7)', () => {
	it('falls back to the built-in when no context can be built', () => {
		const injector = Injector.create({
			parent: TestBed.inject(Injector),
			providers: [
				{ provide: ViewerCustomizationService, deps: [] },
				{ provide: ViewerRegionService, deps: [] },
			],
		});
		injector.get(ViewerCustomizationService).api.setRegion('notes', 'x-notes');
		const regions = injector.get(ViewerRegionService);
		// Not attached yet: no state to build a context from.
		expect(regions.renderer('notes')).toBeUndefined();
		regions.attach(baseState({ selection: [] }));
		expect(regions.renderer('notes')).toBe('x-notes');
		// No options store: the Options region has no context, so no renderer.
		injector.get(ViewerCustomizationService).api.setRegion('dialog.options', 'x-options');
		const dialog = runInInjectionContext(injector, () => new SettingsDialogComponent());
		const view = dialog as unknown as { regionRenderer: () => unknown };
		expect(view.regionRenderer()).toBeUndefined();
	});
});

describe('dialog.options region', () => {
	class OptionsElement extends HTMLElement {
		context: ViewerSettingsRegionContext | undefined;
	}

	it('hands the dialog body to a host element with a settings context that respects locks', async () => {
		customElements.get('x-options-view') ?? customElements.define('x-options-view', OptionsElement);
		const injector = Injector.create({
			parent: TestBed.inject(Injector),
			providers: [
				{ provide: ViewerOptionsService, deps: [] },
				{ provide: ViewerCustomizationService, deps: [] },
				{ provide: ViewerRegionService, deps: [] },
			],
		});
		const customization = injector.get(ViewerCustomizationService);
		const options = injector.get(ViewerOptionsService);
		injector.get(ViewerRegionService).attach(baseState({ selection: [] }));
		const dialog = runInInjectionContext(injector, () => new SettingsDialogComponent());
		const view = dialog as unknown as {
			regionRenderer: () => unknown;
			regionContext: () => ViewerSettingsRegionContext | null;
		};
		expect(view.regionRenderer()).toBeUndefined();
		customization.api.setRegion('dialog.options', 'x-options-view');
		expect(view.regionRenderer()).toBe('x-options-view');

		// The region host creates the element and keeps its `context` current.
		const fixture = await mount(ViewerRegionHostComponent, {
			renderer: 'x-options-view',
			context: view.regionContext(),
		});
		fixture.detectChanges();
		const element = (fixture.nativeElement as HTMLElement).querySelector<OptionsElement>(
			'x-options-view',
		);
		expect(element?.context?.region).toBe('dialog.options');

		const ctx = view.regionContext();
		expect(typeof ctx?.setOption).toBe('function');
		ctx?.setOption('advanced.showGrid', true);
		expect(options.options().advanced.showGrid).toBeTruthy();
		customization.api.lockSetting('advanced.showGrid', false);
		ctx?.setOption('advanced.showGrid', true);
		expect(options.options().advanced.showGrid).toBeFalsy();
		ctx?.resetOptions();
	});
});

describe('backstage.extraPages', () => {
	it('lists the host page in the nav and selects it under backstage.host.<id>', () => {
		const injector = Injector.create({
			parent: TestBed.inject(Injector),
			providers: [
				{ provide: ViewerOptionsService, deps: [] },
				{ provide: ViewerCustomizationService, deps: [] },
				{ provide: ViewerRegionService, deps: [] },
			],
		});
		const customization = injector.get(ViewerCustomizationService);
		injector.get(ViewerRegionService).attach(baseState({ selection: [] }));
		const render = vi.fn();
		customization.api.addBackstagePage({ id: 'team', label: 'My Team', render });
		const file = runInInjectionContext(injector, () => new RibbonFileSectionComponent());
		const view = file as unknown as {
			mainNav: () => readonly { id: string; label: string; host?: unknown }[];
			page: { set: (id: string) => void };
			currentPage: () => string;
			pageRenderer: () => unknown;
			pageContext: () => ViewerRegionContext | null;
		};
		const entry = view.mainNav().find((item) => item.id === 'host.team');
		expect(entry?.label).toBe('My Team');
		expect(view.pageRenderer()).toBeUndefined();
		view.page.set('host.team');
		expect(view.currentPage()).toBe('host.team');
		expect(view.pageRenderer()).toBe(render);
		expect(view.pageContext()?.region).toBe('backstage.host.team');
		customization.api.removeBackstagePage('team');
		expect(view.currentPage()).toBe('home');
	});
});

describe('capabilities.edit', () => {
	it('flips editingAllowed live and restores when cleared', async () => {
		await mount(ViewerRegionHostComponent, { renderer: 'x', context: context([]) });
		const customization = TestBed.inject(ViewerCustomizationService);
		expect(customization.editingAllowed()).toBeTruthy();
		customization.api.setCapability('edit', false);
		expect(customization.editingAllowed()).toBeFalsy();
		customization.api.setCapability('edit', undefined);
		expect(customization.editingAllowed()).toBeTruthy();
	});
});
