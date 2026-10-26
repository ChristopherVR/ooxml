// @vitest-environment happy-dom
import { PptxHandler } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	embedFontsOnLoad,
	retainedFontLists,
} from '../../../../shared/test-utils/outline-fonts.mjs';
import {
	presentationLoadFixtures,
	settlePresentationLoad,
} from '../../../../shared/test-utils/presentation-load.mjs';
import { glyphOutlineFontCache } from '../utils/glyph-outline-cache';
import type { EditorHistoryResult } from './useEditorHistory';
import { useLoadContent } from './useLoadContent';
import { useViewerState } from './useViewerState';

afterEach(() => vi.restoreAllMocks());

describe('embedded fonts in the glyph-outline cache', () => {
	it('are released when the deck is replaced and when the viewer unmounts', async () => {
		const fixtures = await presentationLoadFixtures(PptxHandler);
		const lists = embedFontsOnLoad(PptxHandler);
		let state!: ReturnType<typeof useViewerState>;
		function Harness({ content }: { content: Uint8Array }) {
			state = useViewerState({ content, canEdit: true });
			useLoadContent({
				...state,
				content,
				clearSelection: vi.fn(),
				history: { resetHistory: vi.fn() } as unknown as EditorHistoryResult,
				setIsEncrypted: vi.fn(),
				setReadOnlyRecommendation: vi.fn(),
				setModifyVerifier: vi.fn(),
				setCompatToasts: vi.fn(),
			});
			return null;
		}
		const host = document.createElement('div');
		const root = createRoot(host);
		const loaded = async (title: string) => {
			await vi.waitFor(async () => {
				await act(settlePresentationLoad);
				expect(state.coreProperties?.title).toBe(title);
			});
		};
		try {
			await act(async () => {
				root.render(<Harness content={fixtures.first} />);
			});
			await loaded('Presentation A');
			expect(state.embeddedFonts).toBe(lists[0]);
			expect(retainedFontLists(glyphOutlineFontCache).has(lists[0]!)).toBeTruthy();

			await act(async () => {
				root.render(<Harness content={fixtures.second} />);
			});
			await loaded('Presentation B');
			const retained = retainedFontLists(glyphOutlineFontCache);
			expect(retained.has(lists[0]!)).toBeFalsy();
			expect(retained.has(lists[1]!)).toBeTruthy();
		} finally {
			await act(async () => root.unmount());
			host.remove();
		}
		const retained = retainedFontLists(glyphOutlineFontCache);
		expect(lists).toHaveLength(2);
		expect(lists.some((list) => retained.has(list))).toBeFalsy();
	}, 30_000);
});
