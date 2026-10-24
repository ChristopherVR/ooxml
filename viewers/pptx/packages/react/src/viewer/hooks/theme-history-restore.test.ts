import { PRESET_THEMES, settleThemeWrites } from 'ooxml-ui/pptx';
import type { PptxHandler } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import type { EditorHistorySnapshot } from '../types';
import { serializeHistoryDocument } from './explicit-slide-history';
import { restoreHistoryTheme } from './theme-history-restore';
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

describe('the change gate', () => {
	it('sees a theme change, so a pick that leaves the slides alone is still an undo step', () => {
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

	it('writes on the shared theme write queue, which every save waits for', async () => {
		const { input, handler, writes } = harness({
			colorScheme: office.colorScheme,
			fontScheme: arial,
		});
		handler.updateThemeFontScheme.mockImplementation(async () => {
			await new Promise<void>((done) => {
				setTimeout(done, 20);
			});
			writes.push('fonts');
		});
		void restoreHistoryTheme(input, { colorScheme: office.colorScheme, fontScheme: calibri });
		await settleThemeWrites(input.handlerRef.current);
		expect(writes).toStrictEqual(['fonts']);
	});
});
