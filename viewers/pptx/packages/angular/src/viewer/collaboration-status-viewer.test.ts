/**
 * A keyless peer-to-peer link says why it cannot join, on the page.
 *
 * The viewer mounts on a webrtc session config without its secret (what opening
 * `?room=...&transport=webrtc` with no `#key=` gives a demo), and the status bar shows
 * the same pill the other bindings show: "Connection error", a Retry button, and the
 * reason as the pill's tooltip, without opening Share.
 */
import { TestBed } from '@angular/core/testing';
import type { ComponentFixture } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { PowerPointViewerComponent } from './power-point-viewer.component';
import { declareSignalFields } from './signal-inputs.test-support';

beforeAll(async () => {
	declareSignalFields(PowerPointViewerComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

afterEach(() => TestBed.resetTestingModule());

async function settle(fixture: ComponentFixture<unknown>, done: () => boolean): Promise<void> {
	for (let attempt = 0; attempt < 100 && !done(); attempt += 1) {
		fixture.detectChanges();
		await fixture.whenStable();
		await new Promise((resolve) => {
			setTimeout(resolve, 20);
		});
	}
	fixture.detectChanges();
}

describe('collaboration status (Angular viewer)', () => {
	it('shows a keyless link as a visible error with Retry and the reason', async () => {
		const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
		const content = await handler.save(data.slides);
		handler.dispose();
		TestBed.configureTestingModule({ providers: [provideTranslateService()] });
		const translate = TestBed.inject(TranslateService);
		translate.setTranslation('en', translationsEn);
		translate.use('en');
		const fixture = TestBed.createComponent(PowerPointViewerComponent);
		fixture.componentRef.setInput('content', content);
		fixture.componentRef.setInput('collaboration', {
			roomId: 'keyless-room',
			serverUrl: '',
			userName: 'Ada',
			transport: 'webrtc',
		});
		const root = fixture.nativeElement as HTMLElement;
		const pill = () =>
			root.querySelector<HTMLElement>(
				'[slot="collaboration"] [data-testid="collaboration-status"]',
			);
		await settle(fixture, () => pill() !== null);

		const status = pill();
		expect(status).not.toBeNull();
		expect(status!.textContent).toContain(translationsEn['pptx.collaboration.status.error']);
		expect(status!.getAttribute('title')).toBe(
			translationsEn['pptx.collaboration.failure.missingSecret'],
		);
		const retry = status!.querySelector('button');
		expect(retry?.textContent?.trim()).toBe(translationsEn['pptx.collaboration.retry']);
		// One status region, so the reason is not announced twice.
		expect(root.querySelectorAll('[data-testid="collaboration-status"]')).toHaveLength(1);
		fixture.destroy();
	});
});
