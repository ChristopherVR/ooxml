import { expect, it } from 'vitest';
import { visioIsoDateSerial } from './date-cache';
import { visioSerialDate } from './text-fields';

it('retains native wall-clock ISO date caches independent of local timezone', () => {
	for (const text of [
		'2026-10-13',
		'2026-10-13T14:15',
		'2026-10-13T23:59:59.999',
		'2024-02-29T00:00:00.001',
	]) {
		const date = visioSerialDate(visioIsoDateSerial(text)!);
		expect(date?.toISOString()).toBe(
			text.length === 10
				? text + 'T00:00:00.000Z'
				: text.length === 16
					? text + ':00.000Z'
					: text + 'Z',
		);
	}
});
it('keeps invalid Gregorian dates and zoned caches unsupported instead of normalizing them', () => {
	for (const text of [
		'2026-02-30',
		'1900-02-29',
		'0000-01-01',
		'2026-13-01',
		'2026-01-00',
		'2026-01-01T24:00:00',
		'2026-01-01T00:60:00',
		'2026-01-01T00:00:60',
		'2026-01-01T00:00:00Z',
		'30177-04-13',
	])
		expect(visioIsoDateSerial(text)).toBeUndefined();
});
