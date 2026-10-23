import { PRESET_THEMES } from 'ooxml-ui/pptx';
import type { PptxHandler } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import type { EditorHistorySnapshot } from '../types';
import { serializeHistoryDocument } from './explicit-slide-history';
import { historyTheme, withHistoryTheme } from './history-theme';
import { queueThemeWrite, restoreHistoryTheme } from './theme-history-restore';
import type { UseThemeHandlersInput } from './useThemeHandlers';

const office = PRESET_THEMES[0];
const calibri = { majorFont: { latin: 'Calibri Light' }, minorFont: { latin: 'Calibri' } };
const arial = { majorFont: { latin: 'Arial' }, minorFont: { latin: 'Arial' } };

function harness(theme: UseThemeHandlersInput['theme']) {
	const writes: string[] = [];
	const handler = {
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
	let state: Record<string, unknown> | null = { ...theme, formatScheme: { id: 'kept' } };
	const input = {
		handlerRef: { current: handler as unknown as PptxHandler },
		theme,
		setTheme: vi.fn((update) => {
			state = typeof update === 'function' ? update(state) : update;
		}),
		slideMasters: [{ themePath: 'ppt/theme/theme1.xml' }],
		themeOptions: [
			{ path: 'ppt/theme/theme2.xml', name: 'Other' },
			{ path: 'ppt/theme/theme1.xml', name: 'Office Theme' },
		],
	} as unknown as UseThemeHandlersInput;
	return { input, handler, writes, state: () => state };
}

describe('history theme records', () => {
	it('records the schemes and the name in one order, and the change gate sees them', () => {
		const recorded = historyTheme({
			name: 'Facet',
			fontScheme: arial,
			colorScheme: office.colorScheme,
		});
		expect(Object.keys(recorded ?? {})).toStrictEqual(['colorScheme', 'fontScheme', 'name']);
		const snapshot = {
			width: 1,
			height: 1,
			activeSlideIndex: 0,
			slides: [],
			templateElementsBySlideId: {},
		} as EditorHistorySnapshot;
		expect(serializeHistoryDocument({ ...snapshot, theme: { fontScheme: arial } })).not.toBe(
			serializeHistoryDocument({ ...snapshot, theme: { fontScheme: calibri } }),
		);
		expect(serializeHistoryDocument(snapshot)).toBe(
			serializeHistoryDocument({ ...snapshot, theme: undefined }),
		);
	});

	it('restores exactly the recorded parts and keeps the rest of the theme', () => {
		const restored = withHistoryTheme(
			{ name: 'Facet', fontScheme: arial, colorScheme: office.colorScheme, formatScheme: 'kept' },
			{ colorScheme: office.colorScheme, fontScheme: calibri },
		);
		expect(restored).toStrictEqual({
			colorScheme: office.colorScheme,
			fontScheme: calibri,
			formatScheme: 'kept',
		});
	});
});

describe('restoreHistoryTheme', () => {
	it('undoing a fonts pick writes only the font scheme back', async () => {
		const { input, handler, writes, state } = harness({
			colorScheme: office.colorScheme,
			fontScheme: arial,
		});
		await restoreHistoryTheme(input, { colorScheme: office.colorScheme, fontScheme: calibri });
		expect(writes).toStrictEqual(['fonts']);
		expect(handler.updateThemeFontScheme).toHaveBeenCalledWith(calibri);
		expect(state()).toStrictEqual({
			colorScheme: office.colorScheme,
			fontScheme: calibri,
			formatScheme: { id: 'kept' },
		});
	});

	it('undoing back to the loaded, nameless theme puts the primary part name back', async () => {
		const { input, writes, state } = harness({
			colorScheme: { ...office.colorScheme, accent1: '#90C226' },
			fontScheme: arial,
			name: 'Facet',
		});
		await restoreHistoryTheme(input, { colorScheme: office.colorScheme, fontScheme: calibri });
		expect(writes).toStrictEqual(['colors', 'fonts', 'name:Office Theme']);
		expect(state()).not.toHaveProperty('name');
	});

	it('runs theme part writes one after another', async () => {
		const { handler } = harness(undefined);
		const order: string[] = [];
		let release: () => void = () => {};
		const slow = new Promise<void>((done) => {
			release = done;
		});
		const first = queueThemeWrite(handler as unknown as PptxHandler, async () => {
			await slow;
			order.push('first');
		});
		const second = queueThemeWrite(handler as unknown as PptxHandler, async () => {
			order.push('second');
		});
		release();
		await Promise.all([first, second]);
		expect(order).toStrictEqual(['first', 'second']);
	});
});
