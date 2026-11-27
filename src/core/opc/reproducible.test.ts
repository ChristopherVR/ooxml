import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	OFFICE_ZIP_DATE,
	dateZipEntries,
	guidFrom,
	idInRange,
	randomIdSource,
	saveEnvironment,
	seededIdSource,
} from './reproducible';

const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/u;

describe('reproducible save options', () => {
	it('gives the same sequence for the same seed and another for another seed', () => {
		const take = (seed: number) => {
			const ids = seededIdSource(seed);
			return Array.from({ length: 5 }, () => ids.next());
		};
		expect(take(7)).toEqual(take(7));
		expect(take(7)).not.toEqual(take(8));
		for (const value of take(7)) expect(value >>> 0).toBe(value);
	});

	it('makes version 4 GUIDs whose every field varies', () => {
		const ids = seededIdSource(1);
		const a = guidFrom(ids);
		const b = guidFrom(ids);
		expect(a).toMatch(GUID);
		expect(b).toMatch(GUID);
		expect(a.slice(1, 9)).not.toBe(b.slice(1, 9));
		expect(guidFrom(randomIdSource)).toMatch(GUID);
	});

	it('draws integers inside the range', () => {
		const ids = seededIdSource(3);
		for (let i = 0; i < 100; i++) {
			const value = idInRange(ids, 1, 2 ** 31 - 1);
			expect(value).toBeGreaterThanOrEqual(1);
			expect(value).toBeLessThanOrEqual(2 ** 31 - 1);
		}
	});

	it('fills in the defaults and rejects an invalid date', () => {
		const env = saveEnvironment();
		expect(env.ids).toBe(randomIdSource);
		expect(env.zipDate).toBe(OFFICE_ZIP_DATE);
		expect(Math.abs(env.now.getTime() - Date.now())).toBeLessThan(5000);
		const now = new Date('2026-01-15T09:00:00Z');
		expect(saveEnvironment({ now }).now).toBe(now);
		expect(() => saveEnvironment({ now: new Date('nope') })).toThrow(RangeError);
	});

	it('dates every ZIP entry as Office does (1980-01-01)', async () => {
		const zip = new JSZip();
		zip.file('a.xml', '<a/>');
		zip.file('b/c.xml', '<c/>');
		dateZipEntries(zip);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const view = new DataView(bytes.buffer, bytes.byteOffset);
		// The first local header: DOS time at offset 10, DOS date at 12.
		expect(view.getUint16(10, true)).toBe(0);
		expect(view.getUint16(12, true)).toBe(0x21);
	});
});

describe('bindSaveEnvironment', () => {
	it('gives the writers of one save its environment, and defaults to a fresh one', async () => {
		const { bindSaveEnvironment, saveEnvironmentOf } = await import('./reproducible');
		const zip = new JSZip();
		const env = saveEnvironment({ now: new Date('2026-01-15T09:00:00Z'), ids: seededIdSource(1) });
		expect(bindSaveEnvironment(zip, env)).toBe(zip);
		expect(saveEnvironmentOf(zip)).toBe(env);
		const other = new JSZip();
		expect(saveEnvironmentOf(other).ids).toBe(randomIdSource);
		expect(saveEnvironmentOf(other)).toBe(saveEnvironmentOf(other));
	});
});
