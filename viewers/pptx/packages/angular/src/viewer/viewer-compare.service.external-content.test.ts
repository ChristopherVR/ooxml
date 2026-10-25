// @vitest-environment happy-dom
import { Injector, signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { TranslateService } from '@ngx-translate/core';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { ViewerCompareService } from './viewer-compare.service';
import { ViewerDialogsService } from './viewer-dialogs.service';
import { ViewerOptionsService } from './viewer-options.service';

beforeAll(() => {
	try {
		TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	} catch {
		// Already initialised by another spec in this worker.
	}
});

/** Answer the next file picker with a deck and capture the options Compare loads it with. */
function pickDeckAndCaptureLoad() {
	vi.spyOn(HTMLInputElement.prototype, 'click').mockImplementation(function (
		this: HTMLInputElement,
	) {
		Object.defineProperty(this, 'files', { value: [new File(['pptx'], 'other.pptx')] });
		this.dispatchEvent(new Event('change'));
	});
	return vi.spyOn(PptxHandler.prototype, 'load').mockResolvedValue({ slides: [] } as never);
}

function compareService(loader?: Pick<LoadContentService, 'resolveAllowExternalImages'>) {
	const injector = Injector.create({
		parent: TestBed.inject(Injector),
		providers: [
			{ provide: ViewerOptionsService, deps: [] },
			{
				provide: ViewerDialogsService,
				useValue: { compareResult: signal(null), showCompare: signal(false) },
			},
			{ provide: EditorStateService, useValue: { slides: signal([]) } },
			{ provide: TranslateService, useValue: { instant: (key: string) => key } },
			...(loader ? [{ provide: LoadContentService, useValue: loader }] : []),
			{ provide: ViewerCompareService, deps: [] },
		],
	});
	return {
		service: injector.get(ViewerCompareService),
		options: injector.get(ViewerOptionsService),
	};
}

describe('viewerCompareService linked content', () => {
	afterEach(() => {
		vi.restoreAllMocks();
		localStorage.clear();
	});

	it('loads the incoming deck with the same rule as the main load', async () => {
		const load = pickDeckAndCaptureLoad();
		// The host switch says true while the Trust Center option is off.
		const { service, options } = compareService({ resolveAllowExternalImages: () => true });
		options.setValue('trust', 'allowExternalContent', false);
		service.startCompare();
		await vi.waitFor(() => expect(load).toHaveBeenCalled());
		expect(load.mock.calls[0][1]).toStrictEqual({ allowExternalImages: true });
	});

	it('falls back to the Trust Center option without a loader', async () => {
		const load = pickDeckAndCaptureLoad();
		const { service, options } = compareService();
		options.setValue('trust', 'allowExternalContent', false);
		service.startCompare();
		await vi.waitFor(() => expect(load).toHaveBeenCalled());
		expect(load.mock.calls[0][1]).toStrictEqual({ allowExternalImages: false });
	});
});
