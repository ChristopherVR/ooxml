import type { DocumentSaveMeta } from 'ooxml-core/embedding';
import { describe, expect, it, vi } from 'vitest';

import { statusBarState } from '../status-bar-state';
import {
	createSaveProgressController,
	saveProgressStatus,
	withSaveProgress,
} from './save-progress-controller';

describe('save progress controller', () => {
	it('shows each step while a save runs and clears when it resolves', async () => {
		const controller = createSaveProgressController();
		const seen: (string | undefined)[] = [];
		controller.subscribe(() => seen.push(saveProgressStatus(controller.getSnapshot())?.text));
		await controller.track(async (onProgress) => {
			onProgress({ phase: 'checking' });
			onProgress({ phase: 'uploading', loaded: 45, total: 100 });
			onProgress({ phase: 'merging' });
		});
		expect(seen).toEqual(['Checking for changes…', 'Saving… 45%', 'Combining edits…', undefined]);
	});

	it('clears when the save rejects and passes the error on', async () => {
		const controller = createSaveProgressController();
		await expect(
			controller.track(async (onProgress) => {
				onProgress({ phase: 'checking' });
				throw new Error('offline');
			}),
		).rejects.toThrow('offline');
		expect(controller.getSnapshot().progress).toBeUndefined();
	});

	it('announces only phase changes and 10% steps', () => {
		const controller = createSaveProgressController();
		controller.report({ phase: 'uploading', loaded: 11, total: 100 });
		controller.report({ phase: 'uploading', loaded: 14, total: 100 });
		expect(saveProgressStatus(controller.getSnapshot())).toEqual({
			text: 'Saving… 14%',
			announcement: 'Saving… 11%',
		});
		controller.report({ phase: 'uploading', loaded: 21, total: 100 });
		expect(saveProgressStatus(controller.getSnapshot())?.announcement).toBe('Saving… 21%');
	});

	it('translates through the keys and falls back to English for a missing one', () => {
		const controller = createSaveProgressController();
		controller.report({ phase: 'uploading', loaded: 1, total: 2 });
		const t = vi.fn((key: string) =>
			key === 'pptx.autosave.progress.uploading' ? 'Speichern… {percent}%' : key,
		);
		expect(saveProgressStatus(controller.getSnapshot(), t)?.text).toBe('Speichern… 50%');
		controller.report({ phase: 'merging' });
		expect(saveProgressStatus(controller.getSnapshot(), t)?.text).toBe('Combining edits…');
	});

	it('wraps a storage so its saves report progress and other members pass through', async () => {
		const controller = createSaveProgressController();
		const callerSeen: string[] = [];
		const storage = {
			label: 'host',
			async save(_bytes: Uint8Array, meta: DocumentSaveMeta) {
				meta.onProgress?.({ phase: 'uploading', loaded: 1, total: 4 });
				expect(saveProgressStatus(controller.getSnapshot())?.text).toBe('Saving… 25%');
				return { fileName: 'a.pptx' };
			},
		};
		const wrapped = withSaveProgress(storage, controller);
		expect(wrapped.label).toBe('host');
		const result = await wrapped.save(new Uint8Array(), {
			fileName: 'a.pptx',
			mimeType: 'x',
			reason: 'save',
			onProgress: (p) => callerSeen.push(p.phase),
		});
		expect(result).toEqual({ fileName: 'a.pptx' });
		expect(callerSeen).toEqual(['uploading']);
		expect(controller.getSnapshot().progress).toBeUndefined();
	});

	it('replaces the save indicator in the status bar and falls back afterwards', () => {
		const translate = (key: string) => key;
		const base = { slideCount: 1, activeSlideIndex: 0, dirty: false, translate };
		const progress = { text: 'Saving… 5%', announcement: 'Saving… 5%' };
		const during = statusBarState({ ...base, saveProgress: progress });
		expect(during.saveText).toBe('Saving… 5%');
		expect(during.saveKind).toBe('saving');
		expect(during.saveAnnouncement).toBe('Saving… 5%');
		const after = statusBarState(base);
		expect(after.saveText).toBe('pptx.statusBar.allSaved');
		expect(after.saveAnnouncement).toBeUndefined();
	});
});
