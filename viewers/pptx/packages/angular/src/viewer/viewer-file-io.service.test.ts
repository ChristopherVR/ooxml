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
		storage: () => storage,
		fileName: () => 'Quarterly.pptx',
	});
	return { service, savePresentation, afterSuccessfulSave };
}

describe('ViewerFileIOService host storage', () => {
	it('saves to storage.save instead of downloading', async () => {
		const save = vi.fn(async () => undefined);
		const { service, savePresentation, afterSuccessfulSave } = harness({ save });
		await service.saveAs('pptx');
		expect(save).toHaveBeenCalledExactlyOnceWith(bytes, {
			fileName: 'Quarterly.pptx',
			mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			reason: 'save',
		});
		expect(savePresentation).not.toHaveBeenCalled();
		expect(afterSuccessfulSave).toHaveBeenCalledWith('pptx');
	});

	it('does not run the recovery cleanup when the host save fails', async () => {
		const { service, afterSuccessfulSave } = harness({
			save: () => Promise.reject(new Error('offline')),
		});
		await expect(service.saveAs('pptx')).rejects.toThrow('offline');
		expect(afterSuccessfulSave).not.toHaveBeenCalled();
	});

	it('downloads under the host file name without a storage', async () => {
		const { service, savePresentation } = harness(undefined);
		await service.saveAs('ppsx');
		expect(savePresentation).toHaveBeenCalledWith(bytes, 'Quarterly.ppsx', 'ppsx');
	});
});
