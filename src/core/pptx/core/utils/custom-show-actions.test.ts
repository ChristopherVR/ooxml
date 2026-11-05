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

describe('custom show text-link mirrors', () => {
	const jump = 'ppaction://customshow?id=7&return=true';
	const textSlide = (
		textStyle: Record<string, unknown>,
		styles: Record<string, unknown>[],
	): PptxSlide =>
		({
			rId: 'r',
			elements: [
				{
					id: 'text',
					type: 'text',
					textStyle,
					textSegments: styles.map((style) => ({ text: 'x', style })),
				},
			],
		}) as unknown as PptxSlide;
	const clear = (id: string) => (id === '7' ? null : id);

	it('clears run links and their default style without counting the mirror twice', () => {
		const linked = {
			hyperlink: jump,
			hyperlinkAction: jump,
			hyperlinkTooltip: 'show',
			hyperlinkRId: '',
			hyperlinkSoundXml: { '@_r:embed': 'sound' },
			bold: true,
		};
		const external = { hyperlink: 'https://example.com', italic: true };
		const input = textSlide(linked, [linked, external]);
		const result = mapCustomShowActions([input], clear);
		const text = result.slides[0]!.elements[0]!;
		expect(result.changed).toBe(1);
		expect(text.textStyle?.hyperlink).toBeUndefined();
		expect(text.textStyle?.hyperlinkSoundXml).toBeUndefined();
		expect(text.textSegments?.[0]?.style.hyperlink).toBeUndefined();
		expect(text.textSegments?.[1]?.style).toBe(external);
		expect(input.elements[0]!.textStyle).toBe(linked);
		expect(linked.hyperlink).toBe(jump);
	});

	it.each([true, false])(
		'rewrites the canonical link with optional action metadata (%s)',
		(withAction) => {
			const linked = {
				hyperlink: jump,
				...(withAction ? { hyperlinkAction: jump } : {}),
				bold: true,
			};
			const input = textSlide(linked, [linked]);
			const result = mapCustomShowActions([input], () => '12');
			expect(result.changed).toBe(1);
			expect(result.slides[0]!.elements[0]!.textStyle?.hyperlink).toBe(
				'ppaction://customshow?id=12&return=true',
			);
			expect(result.slides[0]!.elements[0]!.textSegments?.[0]?.style.hyperlink).toBe(
				'ppaction://customshow?id=12&return=true',
			);
			expect(linked.hyperlink).toBe(jump);
		},
	);

	it('handles a default-only text link and leaves unrelated links identical', () => {
		const input = textSlide({ hyperlink: jump }, []);
		expect(mapCustomShowActions([input], clear).changed).toBe(1);
		const external = textSlide({ hyperlink: 'https://example.com' }, []);
		const result = mapCustomShowActions([external], clear);
		expect(result.changed).toBe(0);
		expect(result.slides[0]).toBe(external);
	});

	it('ignores an external URL containing the action text and malformed percent escapes', () => {
		expect(
			customShowActionId('https://example.com/?next=ppaction://customshow?id=7'),
		).toBeUndefined();
		expect(customShowActionId('ppaction://customshowcase?id=7')).toBeUndefined();
		expect(customShowActionId('ppaction://customshow?id=%E0%A4%A')).toBeUndefined();
	});
});
