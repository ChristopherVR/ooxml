/**
 * Escape closes the quick-access row's drop-down menus (Presentation options, More actions) and
 * gives focus back to the button that opened them, through the shared `listenForPopupMenuDismiss`.
 */
import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import { RibbonPrimaryRowComponent } from './ribbon-primary-row.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
});
afterEach(() => {
	TestBed.resetTestingModule();
	document.body.replaceChildren();
});

function open() {
	TestBed.configureTestingModule({
		imports: [RibbonPrimaryRowComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	TestBed.overrideComponent(RibbonPrimaryRowComponent, {
		add: { inputs: ['canEdit', 'slideCount'] },
	});
	const fixture = TestBed.createComponent(RibbonPrimaryRowComponent);
	fixture.componentRef.setInput('canEdit', signal(true));
	fixture.componentRef.setInput('slideCount', signal(1));
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.use('en');
	document.body.append(fixture.nativeElement as HTMLElement);
	fixture.detectChanges();
	return fixture;
}

describe('ribbon primary row menus', () => {
	it.each(['Presentation options', 'More actions'])(
		'closes %s on Escape and gives focus back to its button',
		(label) => {
			const fixture = open();
			const root = fixture.nativeElement as HTMLElement;
			const trigger = root.querySelector<HTMLButtonElement>(`[aria-label="${label}"]`)!;
			expect(trigger.getAttribute('aria-haspopup')).toBe('menu');
			trigger.click();
			TestBed.tick();
			expect(trigger.getAttribute('aria-expanded')).toBe('true');
			const escape = new KeyboardEvent('keydown', {
				key: 'Escape',
				bubbles: true,
				cancelable: true,
			});
			document.body.dispatchEvent(escape);
			fixture.detectChanges();
			expect(escape.defaultPrevented).toBe(true);
			expect(trigger.getAttribute('aria-expanded')).toBe('false');
			expect(document.activeElement).toBe(trigger);
		},
	);
});
