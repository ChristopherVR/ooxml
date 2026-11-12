import type { DocumentStorage, SaveRevisionController, ViewerOptions } from 'ooxml-ui/pptx';
import {
	createSaveRevisionController,
	deleteAutosaveSnapshot,
	shouldDiscardAutosaveOnSuccessfulSave,
} from 'ooxml-ui/pptx';

export interface UseSaveRevisionInput {
	/** The stores a save is built from (the ones autosave watches). */
	changeSources: () => readonly unknown[];
	/** Serialise the deck without clearing the dirty flag. */
	getContent: () => Promise<Uint8Array>;
	/** Clear the unsaved-changes state. */
	markClean: () => void;
	filePath: () => string | undefined;
	storage: () => DocumentStorage | undefined;
	getOptions: () => ViewerOptions | undefined;
}

/**
 * The revision-aware acknowledgement of a host-persisted Save (`getSaveRevision`,
 * `getSaveSnapshot`, `acknowledgeSave`): the shared `createSaveRevisionController` plus this
 * binding's dirty flag and recovery cleanup.
 */
export function useSaveRevision(input: UseSaveRevisionInput): SaveRevisionController {
	return createSaveRevisionController({
		getSources: input.changeSources,
		serialize: input.getContent,
		markClean: input.markClean,
		discardRecovery: () => {
			const filePath = input.filePath();
			const options = input.getOptions();
			if (filePath && (!options || shouldDiscardAutosaveOnSuccessfulSave(options))) {
				void deleteAutosaveSnapshot(filePath, input.storage()).catch(() => {
					// The host save succeeded; recovery cleanup is best-effort here.
				});
			}
		},
	});
}
