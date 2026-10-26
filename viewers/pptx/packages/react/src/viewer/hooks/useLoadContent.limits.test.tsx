// @vitest-environment happy-dom
import type { PptxLoadLimits } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	LOAD_LIMIT_DECK_SLIDES,
	loadLimitDeck,
	OVER_SLIDE_LIMIT_MESSAGE,
} from '../../../../shared/test-utils/load-limits.mjs';
import { PptxI18nProvider } from '../i18n/provider';
import type { EditorHistoryResult } from './useEditorHistory';
import { useLoadContent } from './useLoadContent';
import { useViewerState } from './useViewerState';

afterEach(() => vi.restoreAllMocks());

async function renderLoad(limits: PptxLoadLimits | undefined) {
	const content = await loadLimitDeck(PptxHandler);
	const load = vi.spyOn(PptxHandler.prototype, 'load');
	vi.spyOn(console, 'error').mockImplementation(() => {});
	let state!: ReturnType<typeof useViewerState>;
	function Harness() {
		state = useViewerState({ content, canEdit: false });
		useLoadContent({
			...state,
			content,
			clearSelection: vi.fn(),
			history: { resetHistory: vi.fn() } as unknown as EditorHistoryResult,
			setIsEncrypted: vi.fn(),
			setReadOnlyRecommendation: vi.fn(),
			setModifyVerifier: vi.fn(),
			setCompatToasts: vi.fn(),
			limits,
		});
		return null;
	}
	const host = document.createElement('div');
	const root = createRoot(host);
	await act(async () => {
		root.render(
			<PptxI18nProvider translate={() => undefined} locale='en'>
				<Harness />
			</PptxI18nProvider>,
		);
	});
	await vi.waitFor(() => expect(state.loading).toBeFalsy());
	await act(async () => root.unmount());
	return { state, load };
}

describe('useLoadContent limits', () => {
	it('passes the host limits to PptxHandler.load and shows the localized limit error', async () => {
		const limits = { maxSlides: LOAD_LIMIT_DECK_SLIDES - 1 };
		const { state, load } = await renderLoad(limits);
		expect(load).toHaveBeenCalledWith(expect.any(ArrayBuffer), expect.objectContaining({ limits }));
		expect(state.error).toBe(OVER_SLIDE_LIMIT_MESSAGE);
		expect(state.slides).toHaveLength(0);
	});

	it('loads a deck within the limits', async () => {
		const { state } = await renderLoad({ maxSlides: LOAD_LIMIT_DECK_SLIDES });
		expect(state.error).toBeNull();
		expect(state.slides).toHaveLength(LOAD_LIMIT_DECK_SLIDES);
	});

	it('passes undefined when the host sets no limits, keeping the core defaults', async () => {
		const { state, load } = await renderLoad(undefined);
		expect(load.mock.calls[0]?.[1]).toHaveProperty('limits', undefined);
		expect(state.slides).toHaveLength(LOAD_LIMIT_DECK_SLIDES);
	});
});
