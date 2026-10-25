// @vitest-environment happy-dom
import { createViewerOptionsStore } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { loadPresentation } from './load/load-presentation';
import type { ParityWorkflowHost } from './parity-workflows';
import { createParityWorkflows } from './parity-workflows';

vi.mock('./load/load-presentation', () => ({
	loadPresentation: vi.fn(async () => ({ slides: [], blobUrls: [], handler: { dispose() {} } })),
	revokeBlobUrls: vi.fn(),
}));
vi.mock('./ui/compare-panel', () => ({ openComparePanel: vi.fn() }));

function pickDeck() {
	vi.spyOn(HTMLInputElement.prototype, 'click').mockImplementation(function (
		this: HTMLInputElement,
	) {
		Object.defineProperty(this, 'files', { value: [new File(['pptx'], 'other.pptx')] });
		this.dispatchEvent(new Event('change'));
	});
}

async function compareWith(allowExternalContent: boolean | undefined, userOption: boolean) {
	pickDeck();
	const optionsStore = createViewerOptionsStore({ persist: false });
	optionsStore.setValue('trust', 'allowExternalContent', userOption);
	const host = {
		doc: document,
		root: () => document.body,
		t: (key: string) => key,
		store: { get: () => ({ slides: [] }) },
		editor: {},
		optionsStore,
		allowExternalContent,
	} as unknown as ParityWorkflowHost;
	createParityWorkflows(host).openCompare();
	await vi.waitFor(() => expect(loadPresentation).toHaveBeenCalled());
	return vi.mocked(loadPresentation).mock.calls[0][1];
}

describe('parity workflows Compare linked content', () => {
	afterEach(() => {
		vi.restoreAllMocks();
		vi.mocked(loadPresentation).mockClear();
	});

	it('lets the host switch override the Trust Center option', async () => {
		await expect(compareWith(true, false)).resolves.toStrictEqual({ allowExternalImages: true });
	});

	it('uses the Trust Center option while the host switch is unset', async () => {
		await expect(compareWith(undefined, false)).resolves.toStrictEqual({
			allowExternalImages: false,
		});
	});
});
