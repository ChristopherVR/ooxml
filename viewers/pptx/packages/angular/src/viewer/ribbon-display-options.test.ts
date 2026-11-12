import type { DestroyRef } from '@angular/core';
import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { createRibbonDisplayController, ribbonDisplayLayout } from 'ooxml-ui/pptx';
import type { RibbonDisplayController, RibbonDisplayMode, RibbonDisplayState } from 'ooxml-ui/pptx';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { controllerSignal } from './controller-signal';
import { listenRibbonDisplayGestures } from './ribbon-display-listeners';
import { RibbonDisplayMenuComponent } from './ribbon-display-menu.component';
import { RibbonRevealStripComponent } from './ribbon-reveal-strip.component';
import { RibbonTabListComponent } from './ribbon-tab-list.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
});
afterEach(() => TestBed.resetTestingModule());

/** A controller and its state signal, as the root component wires them. */
function displayHarness(initial: RibbonDisplayMode = 'always') {
	const onChange = vi.fn<(mode: RibbonDisplayMode) => void>();
	const controller: RibbonDisplayController = createRibbonDisplayController({
		initial,
		onChange,
	});
	const state = controllerSignal(controller);
	return { controller, onChange, state };
}

function translator(): TranslateService {
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.use('en');
	return translate;
}

/** Signal inputs are declared as plain inputs here (this JIT runner has no signal-input transform). */
function createMenu(controller: RibbonDisplayController, state: () => RibbonDisplayState) {
	TestBed.configureTestingModule({
		imports: [RibbonDisplayMenuComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	TestBed.overrideComponent(RibbonDisplayMenuComponent, {
		add: { inputs: ['state', 'controller'] },
	});
	translator();
	const fixture = TestBed.createComponent(RibbonDisplayMenuComponent);
	fixture.componentRef.setInput('controller', signal(controller));
	fixture.componentRef.setInput('state', signal(state()));
	return fixture;
}

describe('ribbon display options menu', () => {
	it('lists the three modes in PowerPoint order with the current one checked', () => {
		const { controller, state } = displayHarness('tabs');
		const fixture = createMenu(controller, () => state());
		controller.toggleMenu();
		fixture.componentRef.setInput('state', signal(state()));
		fixture.detectChanges();

		const items = [
			...fixture.nativeElement.querySelectorAll('[role="menuitemradio"]'),
		] as HTMLElement[];
		expect(items.map((item) => item.dataset.ribbonDisplay)).toStrictEqual([
			'fullScreen',
			'tabs',
			'always',
		]);
		expect(items.map((item) => item.getAttribute('aria-checked'))).toStrictEqual([
			'false',
			'true',
			'false',
		]);
		const translate = TestBed.inject(TranslateService);
		expect(items[1]!.textContent!).toContain(translate.instant('pptx.ribbon.display.tabs'));
	});

	it('choosing Show tabs only fires the change notification once with tabs', () => {
		const { controller, onChange, state } = displayHarness('always');
		const fixture = createMenu(controller, () => state());
		controller.toggleMenu();
		fixture.componentRef.setInput('state', signal(state()));
		fixture.detectChanges();

		const tabs = fixture.nativeElement.querySelector(
			'[data-ribbon-display="tabs"]',
		) as HTMLButtonElement;
		tabs.click();
		fixture.detectChanges();

		expect(onChange).toHaveBeenCalledTimes(1);
		expect(onChange).toHaveBeenCalledWith('tabs');
		expect(controller.getSnapshot().layout.commands).toBe('hidden');
		expect(controller.getSnapshot().menuOpen).toBe(false);
	});

	it('does not fire the change notification when the host sets the mode', () => {
		const { controller, onChange } = displayHarness('always');
		controller.setMode('fullScreen');
		expect(onChange).not.toHaveBeenCalled();
		expect(controller.getSnapshot().mode).toBe('fullScreen');
	});
});

describe('ribbon display gestures on the tab row', () => {
	function createTabRow(controller: RibbonDisplayController, state: () => RibbonDisplayState) {
		TestBed.configureTestingModule({
			imports: [RibbonTabListComponent],
			providers: [provideTranslateService({ fallbackLang: 'en' })],
		});
		TestBed.overrideComponent(RibbonTabListComponent, {
			add: { inputs: ['activeTab', 'displayState', 'displayController'] },
		});
		translator();
		const fixture = TestBed.createComponent(RibbonTabListComponent);
		fixture.componentRef.setInput('activeTab', signal('home'));
		fixture.componentRef.setInput('displayController', signal(controller));
		fixture.componentRef.setInput('displayState', signal(state()));
		fixture.detectChanges();
		return fixture;
	}

	it('a double click on a tab other than File toggles the ribbon between docked and tabs', () => {
		const { controller, state } = displayHarness('always');
		const fixture = createTabRow(controller, () => state());
		const tabs = fixture.nativeElement.querySelectorAll('[role="tab"]') as NodeListOf<HTMLElement>;
		const file = tabs[0]!;
		const home = tabs[1]!;

		file.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		expect(controller.getSnapshot().mode).toBe('always');

		home.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		expect(controller.getSnapshot().mode).toBe('tabs');
	});

	it('the menu button opens the menu', () => {
		const { controller, state } = displayHarness('always');
		const fixture = createTabRow(controller, () => state());
		const button = fixture.nativeElement.querySelector(
			'[data-pptx-chrome="ribbon-display-options"]',
		) as HTMLButtonElement;
		button.click();
		expect(controller.getSnapshot().menuOpen).toBe(true);
	});
});

describe('ribbon reveal strip', () => {
	it('a press brings the ribbon back over the slide in full-screen mode', () => {
		const { controller } = displayHarness('fullScreen');
		TestBed.configureTestingModule({
			imports: [RibbonRevealStripComponent],
			providers: [provideTranslateService({ fallbackLang: 'en' })],
		});
		TestBed.overrideComponent(RibbonRevealStripComponent, {
			add: { inputs: ['controller'] },
		});
		translator();
		const fixture = TestBed.createComponent(RibbonRevealStripComponent);
		fixture.componentRef.setInput('controller', signal(controller));
		fixture.detectChanges();

		const strip = fixture.nativeElement.querySelector(
			'[data-pptx-chrome="ribbon-reveal"]',
		) as HTMLButtonElement;
		expect(strip.getAttribute('aria-label')).toBe(
			TestBed.inject(TranslateService).instant('pptx.ribbon.showRibbon'),
		);
		expect(controller.getSnapshot().layout.tabs).toBe(false);
		strip.click();
		expect(controller.getSnapshot().layout).toStrictEqual(ribbonDisplayLayout('fullScreen', true));
	});
});

describe('ribbon display document gestures', () => {
	/** A destroy ref that collects the cleanups, as the viewer's DestroyRef does on destroy. */
	function destroyRef(cleanups: (() => void)[]): DestroyRef {
		return {
			onDestroy: (fn: () => void) => {
				cleanups.push(fn);
				return () => undefined;
			},
		} as unknown as DestroyRef;
	}

	it('a press outside the ribbon puts the peeked commands away, and Escape does too', () => {
		const { controller } = displayHarness('tabs');
		const cleanups: (() => void)[] = [];
		listenRibbonDisplayGestures({
			controller,
			root: () => document.body,
			destroyRef: destroyRef(cleanups),
		});
		controller.tabChosen('home');
		expect(controller.getSnapshot().layout.commands).toBe('overlay');

		document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
		expect(controller.getSnapshot().layout.commands).toBe('hidden');

		controller.tabChosen('home');
		document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(controller.getSnapshot().layout.commands).toBe('hidden');

		cleanups.forEach((cleanup) => cleanup());
	});

	it('removes its listeners on destroy', () => {
		const { controller } = displayHarness('tabs');
		const cleanups: (() => void)[] = [];
		listenRibbonDisplayGestures({
			controller,
			root: () => undefined,
			destroyRef: destroyRef(cleanups),
		});
		cleanups.forEach((cleanup) => cleanup());

		controller.tabChosen('home');
		document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
		expect(controller.getSnapshot().layout.commands).toBe('overlay');
		controller.dismiss();
	});
});
