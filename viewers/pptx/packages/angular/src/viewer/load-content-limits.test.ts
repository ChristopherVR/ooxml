/**
 * load-content-limits.test.ts: the host `limits` input reaches
 * `PptxHandler.load`, and a deck over a limit sets the localized error the
 * viewer's load-error screen shows.
 */
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import {
	LOAD_LIMIT_DECK_SLIDES,
	loadLimitDeck,
	OVER_SLIDE_LIMIT_MESSAGE,
} from '../../../shared/test-utils/load-limits.mjs';
import { LoadContentService } from './load-content.service';
import { ViewerOptionsService } from './viewer-options.service';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
});
afterEach(() => {
	TestBed.resetTestingModule();
	vi.restoreAllMocks();
});

function loader(): LoadContentService {
	TestBed.configureTestingModule({
		providers: [
			provideTranslateService({ fallbackLang: 'en' }),
			ViewerOptionsService,
			LoadContentService,
		],
	});
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.use('en');
	return TestBed.inject(LoadContentService);
}

describe('loadContentService host limits', () => {
	it('passes the bound limits to PptxHandler.load and sets the localized limit error', async () => {
		const service = loader();
		const limits = { maxSlides: LOAD_LIMIT_DECK_SLIDES - 1 };
		service.bindHostLimits(() => limits);
		const deck = await loadLimitDeck(PptxHandler);
		const load = vi.spyOn(PptxHandler.prototype, 'load');
		await service.load(deck);
		expect(load).toHaveBeenCalledWith(expect.any(ArrayBuffer), expect.objectContaining({ limits }));
		expect(service.error()).toBe(OVER_SLIDE_LIMIT_MESSAGE);
		expect(service.slides()).toHaveLength(0);
	});

	it('loads a deck within the limits, and passes undefined when none are bound', async () => {
		const service = loader();
		service.bindHostLimits(() => ({ maxSlides: LOAD_LIMIT_DECK_SLIDES }));
		await service.load(await loadLimitDeck(PptxHandler));
		expect(service.error()).toBeNull();
		expect(service.slides()).toHaveLength(LOAD_LIMIT_DECK_SLIDES);

		TestBed.resetTestingModule();
		const unbound = loader();
		const deck = await loadLimitDeck(PptxHandler);
		const load = vi.spyOn(PptxHandler.prototype, 'load');
		await unbound.load(deck);
		expect(load.mock.calls[0]?.[1]).toHaveProperty('limits', undefined);
		expect(unbound.slides()).toHaveLength(LOAD_LIMIT_DECK_SLIDES);
	});
});
