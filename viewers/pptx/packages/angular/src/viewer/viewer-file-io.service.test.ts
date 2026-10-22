import { Injector } from '@angular/core';
import type { DocumentStorage } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import { ExportService } from './export.service';
import { LoadContentService } from './load-content.service';
import { ViewerFileIOService } from './viewer-file-io.service';

/**
 * Host-owned storage: with a `storage` input, Save / Save As hand the bytes to
 * `storage.save()` (the shared `deliverSavedPresentation`) instead of the
 * ExportService download, and the recovery cleanup waits for it to succeed.
 */
const bytes = new Uint8Array([0x50, 0x4b]);

function harness(storage: DocumentStorage | undefined) {
	const savePresentation = vi.fn();
	const afterSuccessfulSave = vi.fn();
	const onSaveError = vi.fn();
	const injector = Injector.create({
		providers: [
			{ provide: LoadContentService, useValue: { saveSlides: vi.fn(async () => bytes) } },
			{ provide: ExportService, useValue: { savePresentation } },
			{ provide: ViewerFileIOService, useClass: ViewerFileIOService },
		],
	});
	const service = injector.get(ViewerFileIOService);
	service.bind({
		canEdit: () => true,
		content: () => null,
		onOpenFile: () => undefined,
		slides: () => [],
		sections: () => [],
		templateElementsBySlideId: () => ({}),
		emitContentChange: () => {},
		saveIntent: () => ({}),
		afterSuccessfulSave,
		onSaveError,
		storage: () => storage,
		fileName: () => 'Quarterly.pptx',
	});
	return { service, savePresentation, afterSuccessfulSave, onSaveError };
}

describe('ViewerFileIOService host storage', () => {
	it('saves to storage.save instead of downloading', async () => {
		const save = vi.fn(async () => undefined);
		const { service, savePresentation, afterSuccessfulSave } = harness({ save });
		await service.saveAs('pptx');
		expect(save).toHaveBeenCalledExactlyOnceWith(bytes, {
			fileName: 'Quarterly.pptx',
			mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			reason: 'saveAs',
		});
		expect(savePresentation).not.toHaveBeenCalled();
		expect(afterSuccessfulSave).toHaveBeenCalledWith('pptx');
	});

	it('reports a failed host save and skips the recovery cleanup', async () => {
		vi.spyOn(console, 'error').mockImplementation(() => {});
		const failure = new Error('offline');
		const { service, afterSuccessfulSave, onSaveError } = harness({
			save: () => Promise.reject(failure),
		});
		await expect(service.saveAs('pptx')).resolves.toBeUndefined();
		expect(onSaveError).toHaveBeenCalledWith(failure);
		expect(afterSuccessfulSave).not.toHaveBeenCalled();
		vi.restoreAllMocks();
	});

	it('reports Save as save and Save As .pptx as saveAs', async () => {
		const save = vi.fn(async () => undefined);
		const { service } = harness({ save });
		await service.save();
		await service.saveAs('pptx');
		expect(save.mock.calls.map(([, meta]) => (meta as { reason: string }).reason)).toEqual([
			'save',
			'saveAs',
		]);
	});

	it('downloads under the host file name without a storage', async () => {
		const { service, savePresentation } = harness(undefined);
		await service.saveAs('ppsx');
		expect(savePresentation).toHaveBeenCalledWith(bytes, 'Quarterly.ppsx', 'ppsx');
	});
});
