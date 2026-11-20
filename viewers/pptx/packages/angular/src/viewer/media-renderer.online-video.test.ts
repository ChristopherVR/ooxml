/**
 * A linked YouTube or Vimeo video paints the provider's player, with the shared iframe
 * permissions. Angular rejects a bound `allow` or `sandbox` on an iframe at render time, so the
 * component used to throw and paint nothing where every other binding shows the video.
 */
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { ONLINE_VIDEO_IFRAME_ALLOW, ONLINE_VIDEO_IFRAME_SANDBOX } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { MediaRendererComponent } from './media-renderer.component';
import { declareSignalFields } from './signal-inputs.test-support';

beforeAll(async () => {
	declareSignalFields(MediaRendererComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});

afterEach(() => TestBed.resetTestingModule());

const onlineVideo = {
	type: 'media',
	id: 'online-video',
	mediaType: 'video',
	isLinked: true,
	mediaPath: 'https://www.youtube.com/watch?v=dQw4w9WgXcQ',
	x: 0,
	y: 0,
	width: 480,
	height: 270,
} as unknown as PptxElement;

describe('online video (Angular media renderer)', () => {
	it('paints the provider player with the shared iframe permissions', () => {
		TestBed.configureTestingModule({ providers: [provideTranslateService()] });
		const translate = TestBed.inject(TranslateService);
		translate.setTranslation('en', translationsEn);
		translate.use('en');
		const fixture = TestBed.createComponent(MediaRendererComponent);
		fixture.componentRef.setInput('element', onlineVideo);
		fixture.detectChanges();

		const frame = (fixture.nativeElement as HTMLElement).querySelector('iframe');
		expect(frame).not.toBeNull();
		expect(frame!.getAttribute('src')).toBe('https://www.youtube.com/embed/dQw4w9WgXcQ');
		expect(frame!.getAttribute('allow')).toBe(ONLINE_VIDEO_IFRAME_ALLOW);
		expect(frame!.getAttribute('sandbox')).toBe(ONLINE_VIDEO_IFRAME_SANDBOX);
	});
});
