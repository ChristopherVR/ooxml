import { describe, expect, it } from 'vitest';

import { UNTITLED_SLIDE_TITLE, toAppTitleEntries } from './app-properties-counts';

describe('toAppTitleEntries', () => {
	it('records breaks inside a title as spaces, as PowerPoint does', () => {
		expect(
			toAppTitleEntries(['INS\nTRUCTIONS', 'A SPECIAL \nDAY', 'One\r\nTwo', 'Soft\vbreak']),
		).toEqual(['INS TRUCTIONS', 'A SPECIAL  DAY', 'One Two', 'Soft break']);
	});

	it('lists an empty title as the untitled entry', () => {
		expect(toAppTitleEntries(['', 'Title'])).toEqual([UNTITLED_SLIDE_TITLE, 'Title']);
	});
});
