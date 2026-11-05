/**
 * Opening from a File > Open cloud place, wired the way `PowerPointViewerComponent` wires it: the
 * file loads through the normal open path (`contentOverride`, a user load), and while it is open
 * the viewer's storage, name, recovery key and edit right are the file's own (the component's
 * `documentView`, the shared `resolveCloudDocumentView`), until another document is loaded.
 *
 * Built in a throwaway injection context with a `DestroyRef` stub and the real loader, like
 * `save-encryption.test.ts`; no TestBed.
 */
import { DestroyRef, Injector } from '@angular/core';
import type { CloudOpenedFile, DocumentStorage } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { ExportService } from './export.service';
import { LoadContentService } from './load-content.service';
import { ViewerFileIOService } from './viewer-file-io.service';

async function deck(slideCount: number): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: slideCount });
	const bytes = await handler.save(data.slides);
	handler.dispose();
	return bytes;
}

async function cloudFile(storage: DocumentStorage, canEdit = true): Promise<CloudOpenedFile> {
	return {
		entry: { id: 'item-1', name: 'Cloud.pptx', kind: 'file' },
		kind: 'pptx',
		fileName: 'Cloud.pptx',
		bytes: await deck(4),
		canEdit,
		storage,
	} as unknown as CloudOpenedFile;
}

function harness(hostStorage: DocumentStorage, hostCanEdit = true) {
	const destroyRef: Pick<DestroyRef, 'onDestroy'> = { onDestroy: () => () => {} };
	const injector = Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: destroyRef },
			{ provide: LoadContentService, useClass: LoadContentService },
			{ provide: ExportService, useValue: { savePresentation: vi.fn() } },
			{ provide: ViewerFileIOService, useClass: ViewerFileIOService },
		],
	});
	const loader = injector.get(LoadContentService);
	const fileIO = injector.get(ViewerFileIOService);
	// The component's `documentView`: host inputs, or the open cloud file's.
	const view = () =>
		fileIO.documentView({
			storage: hostStorage,
			canEdit: hostCanEdit,
			fileName: 'Host.pptx',
			filePath: 'host/Host.pptx',
		});
	let dirty = true;
	const onSaveError = vi.fn();
	fileIO.bind({
		canEdit: () => view().canEdit,
		content: () => null,
		onOpenFile: () => undefined,
		slides: () => loader.slides(),
		sections: () => [],
		templateElementsBySlideId: () => ({}),
		emitContentChange: () => {},
		saveIntent: () => ({}),
		afterSuccessfulSave: () => {},
		onSaveError,
		storage: () => view().storage,
		hostStorage: () => hostStorage,
		fileName: () => view().fileName,
		changeSources: () => [],
		markSaved: () => {
			dirty = false;
		},
	});
	const load = async () => {
		const content = fileIO.activeContent();
		await loader.load(content);
	};
	return { loader, fileIO, view, load, onSaveError, isDirty: () => dirty };
}

describe('ViewerFileIOService opening from a cloud place', () => {
	it('loads the file and saves it to its own storage, not the host storage', async () => {
		const hostSave = vi.fn(async () => undefined);
		const cloudSave = vi.fn(async () => undefined);
		const { loader, fileIO, view, load, isDirty } = harness({ save: hostSave });
		fileIO.contentOverride.set(await deck(2));
		await load();
		expect(loader.slides()).toHaveLength(2);

		fileIO.openCloudFile(await cloudFile({ save: cloudSave }));
		await load();
		expect(loader.slides()).toHaveLength(4);
		expect(view().fileName).toBe('Cloud.pptx');
		expect(view().filePath).toBe('cloud:item-1');
		// Recent and Account keep listing the host storage's copies.
		expect(fileIO.storage()).toEqual({ save: hostSave });

		await fileIO.save();
		expect(cloudSave).toHaveBeenCalledOnce();
		expect(cloudSave.mock.calls[0]?.[1]).toMatchObject({ fileName: 'Cloud.pptx', reason: 'save' });
		expect(hostSave).not.toHaveBeenCalled();
		expect(isDirty()).toBe(false);
	}, 60_000);

	it('stays unsaved when the cloud save fails', async () => {
		vi.spyOn(console, 'error').mockImplementation(() => {});
		const hostSave = vi.fn(async () => undefined);
		const cloudSave = vi.fn(() => Promise.reject(new Error('conflict')));
		const { fileIO, load, onSaveError, isDirty } = harness({ save: hostSave });
		fileIO.openCloudFile(await cloudFile({ save: cloudSave }));
		await load();
		await fileIO.save();
		expect(cloudSave).toHaveBeenCalledOnce();
		expect(hostSave).not.toHaveBeenCalled();
		expect(onSaveError).toHaveBeenCalledOnce();
		expect(isDirty()).toBe(true);
		vi.restoreAllMocks();
	}, 60_000);

	it('opens a file the service leaves read-only without letting it be edited', async () => {
		const { loader, fileIO, view, load } = harness({ save: vi.fn() });
		expect(view().canEdit).toBe(true);
		fileIO.openCloudFile(await cloudFile({ save: vi.fn() }, false));
		await load();
		expect(loader.slides()).toHaveLength(4);
		expect(view().canEdit).toBe(false);
	}, 60_000);

	it('keeps a read-only host read-only for an editable cloud file', async () => {
		const { fileIO, view } = harness({ save: vi.fn() }, false);
		fileIO.openCloudFile(await cloudFile({ save: vi.fn() }));
		expect(view().canEdit).toBe(false);
	}, 60_000);

	it('saves to the host storage again once another document is loaded', async () => {
		const hostSave = vi.fn(async () => undefined);
		const cloudSave = vi.fn(async () => undefined);
		const { fileIO, view, load } = harness({ save: hostSave });
		fileIO.openCloudFile(await cloudFile({ save: cloudSave }));
		await load();
		fileIO.clearCloudFile();
		expect(view().fileName).toBe('Host.pptx');
		expect(view().canEdit).toBe(true);
		await fileIO.save();
		expect(hostSave).toHaveBeenCalledOnce();
		expect(cloudSave).not.toHaveBeenCalled();
	}, 60_000);
});
