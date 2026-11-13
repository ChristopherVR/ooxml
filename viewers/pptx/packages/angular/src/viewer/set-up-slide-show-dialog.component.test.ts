/**
 * The Set Up Slide Show dialog applies the shared settings rules (`applySlideShowSettingsPatch`):
 * a kiosk show loops, a range starts as the whole deck, a custom show defaults to the first.
 * Instantiated directly in a TestBed injection context, with the signal inputs replaced by plain readers.
 */
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import type { PptxCustomShow } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { SetUpSlideShowDialogComponent } from './set-up-slide-show-dialog.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
});
afterEach(() => {
	TestBed.resetTestingModule();
});

const SHOWS: PptxCustomShow[] = [{ id: '4', name: 'Short', slideRIds: ['rId1'] }];

function create(): SetUpSlideShowDialogComponent {
	const component = TestBed.runInInjectionContext(() => new SetUpSlideShowDialogComponent());
	Object.assign(component, { slideCount: () => 6, customShows: () => SHOWS });
	component.draft.set({});
	return component;
}

describe('angular set up slide show dialog follows the shared rules', () => {
	it('a kiosk show loops continuously', () => {
		const component = create();
		component.update({ showType: 'kiosk' });
		expect(component.draft()).toMatchObject({ showType: 'kiosk', loopContinuously: true });
	});

	it('a range starts as the whole deck', () => {
		const component = create();
		component.update({ showSlidesMode: 'range' });
		expect(component.draft()).toMatchObject({
			showSlidesMode: 'range',
			showSlidesFrom: 1,
			showSlidesTo: 6,
		});
	});

	it('a custom show defaults to the first show', () => {
		const component = create();
		component.update({ showSlidesMode: 'customShow' });
		expect(component.draft()).toMatchObject({
			showSlidesMode: 'customShow',
			showSlidesCustomShowId: '4',
		});
	});
});
