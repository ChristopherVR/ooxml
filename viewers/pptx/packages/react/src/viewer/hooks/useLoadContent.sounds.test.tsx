// @vitest-environment happy-dom
import { resolveDeckSound } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { embeddedSoundDeck } from '../../../../shared/test-utils/embedded-sounds.mjs';
import type { EditorHistoryResult } from './useEditorHistory';
import { useLoadContent } from './useLoadContent';
import { useViewerState } from './useViewerState';

afterEach(() => vi.restoreAllMocks());

async function loadDeck(content: Uint8Array) {
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
		});
		return null;
	}
	const host = document.createElement('div');
	const root = createRoot(host);
	await act(async () => {
		root.render(<Harness />);
	});
	await vi.waitFor(() => expect(state.slides).toHaveLength(2), { timeout: 20_000 });
	const loaded = { slides: state.slides, mediaDataUrls: state.mediaDataUrls };
	await act(async () => root.unmount());
	return loaded;
}

describe('useLoadContent embedded sounds', () => {
	it('lets the slide show play embedded transition and effect sounds before linked content is allowed', async () => {
		const { slides, mediaDataUrls } = await loadDeck(await embeddedSoundDeck(PptxHandler));
		const sources = { mediaUrls: mediaDataUrls, linkedContent: false };
		const transitionSound = slides[1]?.transition?.soundPath;
		const effectSound = slides[1]?.nativeAnimations?.find((effect) => effect.soundPath)?.soundPath;

		expect(transitionSound).toMatch(/^ppt\/media\//u);
		expect(effectSound).toMatch(/^ppt\/media\//u);
		expect(resolveDeckSound(transitionSound, sources)).toMatch(/^blob:/u);
		expect(resolveDeckSound(effectSound, sources)).toMatch(/^blob:/u);
	}, 30_000);
});
