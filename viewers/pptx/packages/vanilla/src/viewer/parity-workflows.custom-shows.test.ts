// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { ParityWorkflowHost } from './parity-workflows';
import { createParityWorkflows } from './parity-workflows';
import { openCustomShowsDialog } from './ui/custom-shows-dialog';

vi.mock('./ui/custom-shows-dialog', () => ({ openCustomShowsDialog: vi.fn() }));

const deck = [
	{ id: 's1', rId: 'rId1', elements: [] },
	{ id: 's2', rId: 'rId2', elements: [] },
	{ id: 's3', rId: 'rId3', elements: [] },
];

function runShow(slideRIds: string[], slides: ReadonlyArray<object> = deck) {
	const goToSlide = vi.fn();
	const enterPresentation = vi.fn(async () => {});
	const host = {
		doc: document,
		t: (key: string) => key,
		store: {
			get: () => ({ slides, customShows: [], activeCustomShowId: null, currentSlide: 0 }),
			set: vi.fn(),
		},
		editor: {},
		goToSlide,
		enterPresentation,
	} as unknown as ParityWorkflowHost;
	createParityWorkflows(host).openCustomShows();
	const options = vi.mocked(openCustomShowsDialog).mock.calls[0][2];
	options.onRun({ id: 'show', name: 'Show', slideRIds });
	return { goToSlide, enterPresentation };
}

describe('parity workflows Custom Shows', () => {
	afterEach(() => {
		vi.mocked(openCustomShowsDialog).mockClear();
	});

	it('starts a show on its own first slide, not its first slide in deck order', () => {
		const { goToSlide, enterPresentation } = runShow(['rId3', 'rId2', 'rId1']);
		expect(goToSlide).toHaveBeenCalledExactlyOnceWith(2);
		expect(enterPresentation).toHaveBeenCalledOnce();
	});

	it('skips a hidden first show slide, as the show navigation does', () => {
		const slides = deck.map((slide) => (slide.id === 's3' ? { ...slide, hidden: true } : slide));
		const { goToSlide } = runShow(['rId3', 'rId2', 'rId1'], slides);
		expect(goToSlide).toHaveBeenCalledExactlyOnceWith(1);
	});
});
