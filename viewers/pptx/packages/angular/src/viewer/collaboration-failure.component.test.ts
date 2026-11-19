/**
 * A peer-to-peer session that cannot start says why: the real
 * `CollaborationService` runs the shared session controller against a webrtc
 * config that lacks its invitation secret, and the Share dialog and the status bar's
 * collaboration pill show the translated reason instead of nothing.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
	CUSTOM_ELEMENTS_SCHEMA,
	DestroyRef,
	Injector,
	signal,
	ɵresolveComponentResources as resolveComponentResources,
} from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslatePipe, TranslateService } from '@ngx-translate/core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { CollaborationStatusIndicatorComponent } from './collaboration-status-indicator.component';
import { CollaborationService } from './collaboration.service';
import { ShareDialogComponent } from './share-dialog.component';

const MISSING_SECRET = translationsEn['pptx.collaboration.failure.missingSecret'];

beforeAll(async () => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
	// The JIT build leaves `templateUrl`/`styleUrl` unresolved; read them from disk.
	await resolveComponentResources(async (url) =>
		readFileSync(resolve(import.meta.dirname, url), 'utf8'),
	);
});
afterEach(() => TestBed.resetTestingModule());

function createService(): CollaborationService {
	return Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: { onDestroy: () => () => undefined } as DestroyRef },
			{ provide: CollaborationService, useClass: CollaborationService },
		],
	}).get(CollaborationService);
}

function render<T>(
	component: new (...args: never[]) => T,
	inputs: Record<string, unknown>,
	templateFile?: string,
): HTMLElement {
	return renderFixture(component, inputs, templateFile).nativeElement as HTMLElement;
}

function renderFixture<T>(
	component: new (...args: never[]) => T,
	inputs: Record<string, unknown>,
	templateFile?: string,
) {
	TestBed.resetTestingModule();
	TestBed.configureTestingModule({
		imports: [component],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	if (templateFile) {
		// An override recompiles from `templateUrl`, so inline the template. The
		// dialog's child components are left unrendered (their signal inputs are
		// invisible to the JIT compiler); what they project still renders.
		TestBed.overrideComponent(component, {
			set: {
				template: readFileSync(resolve(import.meta.dirname, templateFile), 'utf8'),
				templateUrl: undefined,
				styleUrl: undefined,
				imports: [TranslatePipe],
				schemas: [CUSTOM_ELEMENTS_SCHEMA],
			} as never,
		});
	}
	// Signal inputs are plain inputs under the JIT test compiler; pass signals.
	TestBed.overrideComponent(component, { add: { inputs: Object.keys(inputs) } });
	const fixture = TestBed.createComponent(component);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, signal(value));
	}
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.use('en');
	fixture.detectChanges();
	return fixture;
}

describe('collaboration start failure (Angular)', () => {
	it('shows why a peer-to-peer session without its secret could not start', async () => {
		const collab = createService();
		await collab.connect({ roomId: 'room-1', serverUrl: '', userName: 'Ada', transport: 'webrtc' });
		expect(collab.status()).toBe('error');
		expect(collab.failure()).toBe('missing-secret');

		const dialog = render(
			ShareDialogComponent,
			{ open: true, active: collab.active(), failureKey: collab.failureKey() },
			'share-dialog.component.html',
		);
		const banner = dialog.querySelector('[data-testid="collaboration-failure"]');
		expect(banner?.textContent?.trim()).toBe(MISSING_SECRET);

		const pill = render(CollaborationStatusIndicatorComponent, {
			status: collab.status(),
			failure: collab.failure(),
		});
		const status = pill.querySelector('[data-testid="collaboration-status"]');
		expect(status?.getAttribute('title')).toBe(MISSING_SECRET);
		expect(status?.textContent).toContain(translationsEn['pptx.collaboration.status.error']);
		expect(status?.getAttribute('aria-label')).toContain(
			translationsEn['pptx.collaboration.status.error'],
		);
		collab.disconnect();
	});

	it('says why a peer-to-peer invitation without its key cannot join', () => {
		const fixture = renderFixture(
			ShareDialogComponent,
			{ open: true, active: false, failureKey: null },
			'share-dialog.component.html',
		);
		const dialog = fixture.nativeElement as HTMLElement;
		dialog.querySelectorAll<HTMLButtonElement>('[role="tab"]')[1]?.click();
		fixture.detectChanges();
		const field = dialog.querySelector<HTMLInputElement>('#pptx-ng-share-invitation');
		field!.value = 'https://x.test/?room=deck-1&transport=webrtc';
		field!.dispatchEvent(new Event('input'));
		fixture.detectChanges();
		const banner = dialog.querySelector('[data-testid="collaboration-failure"]');
		expect(banner?.textContent?.trim()).toBe(MISSING_SECRET);
	});
});
