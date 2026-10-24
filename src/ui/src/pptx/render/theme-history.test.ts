import type { PptxThemeColorScheme } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { saveDeckWithPassword } from './deck-save-encryption';
import {
	crossedHistoryTheme,
	historyTheme,
	loadedThemeName,
	queueThemeWrite,
	restoreCrossedTheme,
	settleThemeWrites,
	themePartWrites,
	withHistoryTheme,
	writeThemeParts,
} from './theme-history';

const OFFICE: PptxThemeColorScheme = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	hlink: '#0563C1',
	folHlink: '#954F72',
};
const FACET: PptxThemeColorScheme = { ...OFFICE, accent1: '#90C226' };
const CALIBRI = { majorFont: { latin: 'Calibri Light' }, minorFont: { latin: 'Calibri' } };
const ARIAL = { majorFont: { latin: 'Arial' }, minorFont: { latin: 'Arial' } };

function writer() {
	const writes: string[] = [];
	return {
		writes,
		updateThemeColorScheme: vi.fn(async () => {
			writes.push('colors');
		}),
		updateThemeFontScheme: vi.fn(async () => {
			writes.push('fonts');
		}),
		updateThemeName: vi.fn(async (name: string) => {
			writes.push(`name:${name}`);
		}),
	};
}

const later = (ms: number) =>
	new Promise<void>((done) => {
		setTimeout(done, ms);
	});

describe('history theme records', () => {
	it('records the schemes and the name in one order, as a copy', () => {
		const theme = { name: 'Facet', fontScheme: ARIAL, colorScheme: FACET, formatScheme: 'x' };
		const recorded = historyTheme(theme);
		expect(Object.keys(recorded ?? {})).toStrictEqual(['colorScheme', 'fontScheme', 'name']);
		expect(recorded?.colorScheme).not.toBe(FACET);
		expect(historyTheme(undefined)).toBeUndefined();
		expect(historyTheme({ name: '' })).toStrictEqual({});
	});

	it('tells a step that crosses a theme change from one that does not', () => {
		const loaded = historyTheme({ colorScheme: OFFICE, fontScheme: CALIBRI });
		const picked = historyTheme({ colorScheme: FACET, fontScheme: ARIAL, name: 'Facet' });
		expect(
			crossedHistoryTheme(loaded, historyTheme({ fontScheme: CALIBRI, colorScheme: OFFICE })),
		).toBeUndefined();
		expect(crossedHistoryTheme(picked, loaded)).toStrictEqual(loaded);
		expect(crossedHistoryTheme(loaded, loaded)).toBeUndefined();
		// A snapshot from before themes were recorded has nothing to restore.
		expect(crossedHistoryTheme(picked, undefined)).toBeUndefined();
	});

	it('restores exactly the recorded parts and keeps the rest of the theme', () => {
		const restored = withHistoryTheme(
			{ name: 'Facet', fontScheme: ARIAL, colorScheme: FACET, formatScheme: 'kept' },
			{ colorScheme: OFFICE, fontScheme: CALIBRI },
		);
		expect(restored).toStrictEqual({
			colorScheme: OFFICE,
			fontScheme: CALIBRI,
			formatScheme: 'kept',
		});
	});

	it('puts a crossed theme back and queues only the parts that differ', async () => {
		const handler = writer();
		const loaded = historyTheme({ colorScheme: OFFICE, fontScheme: CALIBRI });
		const picked = historyTheme({ colorScheme: OFFICE, fontScheme: ARIAL, name: 'Arial' });
		const restored = restoreCrossedTheme(
			{ colorScheme: OFFICE, fontScheme: ARIAL, name: 'Arial', formatScheme: 'kept' },
			picked,
			loaded,
			handler,
			'Office Theme',
		);
		expect(restored).toStrictEqual({
			colorScheme: OFFICE,
			fontScheme: CALIBRI,
			formatScheme: 'kept',
		});
		await settleThemeWrites(handler);
		expect(handler.writes).toStrictEqual(['fonts', 'name:Office Theme']);
		expect(restoreCrossedTheme({ name: 'x' }, loaded, loaded, handler)).toBeUndefined();
		expect(handler.writes).toHaveLength(2);
	});
});

describe('themePartWrites', () => {
	it('writes only the parts that differ: undoing a fonts pick leaves the colours alone', () => {
		expect(
			themePartWrites(
				{ colorScheme: OFFICE, fontScheme: ARIAL },
				{ colorScheme: OFFICE, fontScheme: CALIBRI },
			),
		).toStrictEqual([{ part: 'fontScheme', fontScheme: CALIBRI }]);
	});

	it('puts the loaded part name back when the target theme carries none', () => {
		const writes = themePartWrites(
			{ colorScheme: FACET, fontScheme: ARIAL, name: 'Facet' },
			{ colorScheme: OFFICE, fontScheme: CALIBRI },
			'Office Theme',
		);
		expect(writes.map((write) => write.part)).toStrictEqual(['colorScheme', 'fontScheme', 'name']);
		expect(writes[2]).toStrictEqual({ part: 'name', name: 'Office Theme' });
		expect(
			themePartWrites({ colorScheme: OFFICE }, { colorScheme: OFFICE }, 'Office Theme'),
		).toStrictEqual([]);
	});

	it('finds the primary theme part name, or the first one', () => {
		const options = [
			{ path: 'ppt/theme/theme2.xml', name: 'Other' },
			{ path: 'ppt/theme/theme1.xml', name: 'Office Theme' },
		];
		const masters = [{ themePath: 'ppt/theme/theme1.xml' }, { themePath: 'ppt/theme/theme2.xml' }];
		expect(loadedThemeName(options, masters)).toBe('Office Theme');
		expect(loadedThemeName(options, [])).toBe('Other');
		expect(loadedThemeName(options, undefined)).toBe('Other');
		expect(loadedThemeName(undefined, masters)).toBeUndefined();
	});
});

describe('the theme part write queue', () => {
	it('writes the differing parts one after another, after every write queued before', async () => {
		const handler = writer();
		const first = queueThemeWrite(handler, async () => {
			await later(20);
			handler.writes.push('pick');
		});
		const restore = writeThemeParts(
			handler,
			{ colorScheme: FACET, fontScheme: ARIAL, name: 'Facet' },
			{ colorScheme: OFFICE, fontScheme: CALIBRI },
			'Office Theme',
		);
		await Promise.all([first, restore]);
		expect(handler.writes).toStrictEqual(['pick', 'colors', 'fonts', 'name:Office Theme']);
		await expect(writeThemeParts(null, {}, { colorScheme: OFFICE })).resolves.toBeUndefined();
	});

	it('settles once every write queued so far has landed, even a failed one', async () => {
		const handler = writer();
		await expect(settleThemeWrites(handler)).resolves.toBeUndefined();
		await expect(settleThemeWrites(null)).resolves.toBeUndefined();
		const failing = queueThemeWrite(handler, async () => {
			await later(20);
			throw new Error('write failed');
		});
		const outcome = failing.then(
			() => 'landed',
			(error: Error) => error.message,
		);
		void queueThemeWrite(handler, async () => {
			handler.writes.push('after the failure');
		});
		await settleThemeWrites(handler);
		expect(handler.writes).toStrictEqual(['after the failure']);
		await expect(outcome).resolves.toBe('write failed');
	});

	it('hands back what the write resolves with', async () => {
		await expect(queueThemeWrite(writer(), async () => 42)).resolves.toBe(42);
	});

	it('makes every save wait for the theme part writes queued before it', async () => {
		const order: string[] = [];
		const serializer = {
			save: vi.fn(async () => {
				order.push('save');
				return new Uint8Array([1]);
			}),
			saveEncrypted: vi.fn(async () => new Uint8Array([2])),
		};
		void queueThemeWrite(serializer, async () => {
			await later(30);
			order.push('theme written');
		});
		await saveDeckWithPassword(serializer as never, [], undefined, null);
		expect(order).toStrictEqual(['theme written', 'save']);
	});
});
