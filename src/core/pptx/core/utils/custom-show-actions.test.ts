import { describe, expect, it } from 'vitest';

import type { PptxSlide } from '../types';
import {
	customShowActionId,
	mapCustomShowActions,
	normalizeCustomShowIds,
} from './custom-show-actions';

describe('normalizeCustomShowIds', () => {
	it('keeps valid unique ids and gives the rest the next free number', () => {
		const { shows, idMap } = normalizeCustomShowIds([
			{ id: '3', name: 'a', slideRIds: [] },
			{ id: 'show-uuid', name: 'b', slideRIds: [] },
			{ id: '3', name: 'dup', slideRIds: [] },
			{ id: '99999999999', name: 'too big', slideRIds: [] },
		]);
		expect(shows.map((s) => s.id)).toEqual(['3', '4', '5', '6']);
		expect([...idMap]).toEqual([
			['show-uuid', '4'],
			['3', '5'],
			['99999999999', '6'],
		]);
	});

	it('changes nothing when every id is valid', () => {
		const input = [{ id: '0', name: 'a', slideRIds: [] }];
		const { shows, idMap } = normalizeCustomShowIds(input);
		expect(shows[0]).toBe(input[0]);
		expect(idMap.size).toBe(0);
	});
});

describe('mapCustomShowActions', () => {
	const slide = (action: string): PptxSlide =>
		({
			rId: 'r',
			elements: [{ id: 'e', type: 'shape', actionClick: { action, tooltip: 't' } }],
		}) as unknown as PptxSlide;

	it('reads the id of a customshow action only', () => {
		expect(customShowActionId('ppaction://customshow?id=7&return=true')).toBe('7');
		expect(customShowActionId('ppaction://hlinkshowjump?jump=nextslide')).toBeUndefined();
	});

	it('rewrites the id in place and keeps the rest of the action', () => {
		const { slides, changed } = mapCustomShowActions(
			[slide('ppaction://customshow?id=a&return=true')],
			() => '12',
		);
		expect(changed).toBe(1);
		const el = slides[0]?.elements[0] as unknown as {
			actionClick: { action: string; tooltip: string };
		};
		expect(el.actionClick).toEqual({
			action: 'ppaction://customshow?id=12&return=true',
			tooltip: 't',
		});
	});
});
