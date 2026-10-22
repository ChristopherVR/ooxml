// @vitest-environment happy-dom
import type { DocumentStorage } from 'ooxml-ui/pptx';
import type { PptxData, PptxSaveFormat } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { computed, effectScope, ref } from 'vue';

// `useExportProgress` reads `useI18n`, which needs a component setup.
vi.mock(import('vue-i18n'), async (importOriginal) => ({
	...(await importOriginal()),
	useI18n: (() => ({ t: (key: string) => key })) as never,
}));

const { useExportWiring } = await import('./useExportWiring');

// `downloadAs` is Save / Save As: the bytes go to the host's `storage` when the
// viewer has one, else to a download whose MIME type follows the format (the
// shared `deliverSavedPresentation`, which every binding uses).

const bytes = new Uint8Array([0x50, 0x4b]);

function wiring(storage?: DocumentStorage, fileName = 'deck.pptx') {
	const scope = effectScope();
	const result = scope.run(() =>
		useExportWiring({
			mergedSlides: ref([]),
			slides: ref([]),
			slideCount: computed(() => 0),
			canvasSize: ref({ width: 960, height: 540 }),
			activeSlideIndex: ref(0),
			saveAs: vi.fn(async (_format: PptxSaveFormat) => bytes),
			fileName: () => fileName,
			getDeckData: () => ({}) as PptxData,
			storage: () => storage,
		}),
	)!;
	return { ...result, stop: () => scope.stop() };
}

afterEach(() => {
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});

describe('useExportWiring downloadAs', () => {
	it('hands the save to storage.save instead of downloading', async () => {
		const click = vi.spyOn(HTMLAnchorElement.prototype, 'click');
		const save = vi.fn(async () => undefined);
		const { downloadAs, stop } = wiring({ save });
		await downloadAs('pptx');
		expect(save).toHaveBeenCalledExactlyOnceWith(bytes, {
			fileName: 'deck.pptx',
			mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			reason: 'save',
		});
		expect(click).not.toHaveBeenCalled();
		stop();
	});

	it('downloads under the source name with the legacy MIME type for .ppt', async () => {
		const createObjectURL = vi.fn((_blob: Blob) => 'blob:deck');
		vi.stubGlobal('URL', { createObjectURL, revokeObjectURL: () => {} });
		const names: string[] = [];
		vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (
			this: HTMLAnchorElement,
		) {
			names.push(this.download);
		});
		const { downloadAs, stop } = wiring();
		await downloadAs('ppt');
		expect(names).toStrictEqual(['deck.ppt']);
		expect(createObjectURL.mock.calls[0]![0].type).toBe('application/vnd.ms-powerpoint');
		stop();
	});
});
