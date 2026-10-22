import { describe, expect, it } from 'vitest';
import type { VisioPage, VisioShape } from '../model';
import {
	VISIO_PRESENCE_MAX_COORDINATE,
	VISIO_PRESENCE_MAX_ID,
	VISIO_PRESENCE_MAX_SELECTION,
	sanitizeVisioPresence,
	visioShapeFrame,
} from './collab-presence';

describe('sanitizeVisioPresence', () => {
	it('keeps a valid payload', () => {
		expect(sanitizeVisioPresence({ page: 'p1', x: 1.5, y: 2, selection: ['3', '4'] })).toEqual({
			page: 'p1',
			x: 1.5,
			y: 2,
			selection: ['3', '4'],
		});
	});
	it('hides a peer without a page and accepts a numeric page', () => {
		expect(sanitizeVisioPresence({})).toBeNull();
		expect(sanitizeVisioPresence({ page: {} })).toBeNull();
		expect(sanitizeVisioPresence({ page: 2 })).toEqual({ page: '2' });
	});
	it('drops a pointer that is not a finite pair and clamps a huge one', () => {
		expect(sanitizeVisioPresence({ page: 'p', x: 1 })).toEqual({ page: 'p' });
		expect(sanitizeVisioPresence({ page: 'p', x: NaN, y: 1 })).toEqual({ page: 'p' });
		expect(sanitizeVisioPresence({ page: 'p', x: '1', y: 1 })).toEqual({ page: 'p' });
		expect(sanitizeVisioPresence({ page: 'p', x: 1e12, y: -1e12 })).toEqual({
			page: 'p',
			x: VISIO_PRESENCE_MAX_COORDINATE,
			y: -VISIO_PRESENCE_MAX_COORDINATE,
		});
	});
	it('bounds, dedupes and filters the selection', () => {
		const many = Array.from({ length: 500 }, (_, i) => String(i));
		const result = sanitizeVisioPresence({ page: 'p', selection: [...many, 1, '', null] });
		expect(result?.selection).toHaveLength(VISIO_PRESENCE_MAX_SELECTION);
		const long = sanitizeVisioPresence({
			page: 'x'.repeat(999),
			selection: ['a', 'a', 'b'.repeat(999)],
		});
		expect(long?.page).toHaveLength(VISIO_PRESENCE_MAX_ID);
		expect(long?.selection).toEqual(['a', 'b'.repeat(VISIO_PRESENCE_MAX_ID)]);
		expect(sanitizeVisioPresence({ page: 'p', selection: 'x' })).toEqual({ page: 'p' });
	});
});

describe('visioShapeFrame', () => {
	const shape = (id: string, transform: VisioShape['transform'], children: VisioShape[] = []) =>
		({ id, width: 2, height: 1, transform, children }) as unknown as VisioShape;
	const page = {
		shapes: [shape('1', [1, 0, 0, 1, 10, 5], [shape('2', [1, 0, 0, 1, 1, 1])])],
	} as unknown as VisioPage;
	it('composes parent transforms', () => {
		expect(visioShapeFrame(page, '2')).toEqual({
			transform: [1, 0, 0, 1, 11, 6],
			width: 2,
			height: 1,
		});
		expect(visioShapeFrame(page, '1')?.transform).toEqual([1, 0, 0, 1, 10, 5]);
		expect(visioShapeFrame(page, 'nope')).toBeUndefined();
	});
});
