import { describe, expect, it, vi } from 'vitest';

import type { DeckApi } from '../editor/deck-api';
import type { ExportingApi } from '../export/exporting-api';
import { runQuickAccessCommand } from './quick-access-commands';
import type { ViewerParityUiState } from './viewer-parity-ui.svelte';

function deps(activeSlide: number) {
	const deck = {
		addSlide: vi.fn(),
		getActiveSlideIndex: () => activeSlide,
	} as unknown as DeckApi & { addSlide: ReturnType<typeof vi.fn> };
	return {
		deck,
		exportingApi: {} as ExportingApi,
		parityUi: {} as ViewerParityUiState,
	};
}

describe('quick access commands', () => {
	it('inserts New Slide after the slide on screen, as the ribbon does, not at the end', () => {
		const input = deps(1);
		runQuickAccessCommand('newSlide', input);
		expect(input.deck.addSlide).toHaveBeenCalledExactlyOnceWith(1);
	});
});
