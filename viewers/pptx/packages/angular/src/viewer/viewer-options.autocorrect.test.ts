import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { ViewerOptionsService } from './viewer-options.service';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
});
afterEach(() => {
	TestBed.resetTestingModule();
});

/** AutoCorrect acts on what was typed since `previous`, never on text already in the box. */
describe('ViewerOptionsService.autoCorrect', () => {
	const existing = 'Fine. https://example.com/a has IDs on monday.';

	it('leaves the existing text alone and corrects only the typed words', () => {
		TestBed.configureTestingModule({ providers: [ViewerOptionsService] });
		const options = TestBed.inject(ViewerOptionsService);
		expect(options.autoCorrect(`${existing} typed monday`, existing)).toBe(
			`${existing} Typed Monday`,
		);
	});

	it('still treats the whole text as typed when there is no previous text', () => {
		TestBed.configureTestingModule({ providers: [ViewerOptionsService] });
		const options = TestBed.inject(ViewerOptionsService);
		expect(options.autoCorrect('HEllo monday')).toBe('Hello Monday');
	});
});
