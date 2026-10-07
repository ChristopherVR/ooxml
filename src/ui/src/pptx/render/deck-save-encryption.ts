// Compatibility exports: document operations live in ooxml-core.
export {
	planDeckSave,
	recoverySnapshotIntent,
	saveDeckWithPassword,
	isEncryptedDeckBytes,
} from 'ooxml-core/pptx/editor/render/deck-save-encryption';
export type {
	DeckSaveSerializer,
	DeckSaveOptions,
	DeckSavePurpose,
	DeckSaveIntent,
	DeckSavePlan,
} from 'ooxml-core/pptx/editor/render/deck-save-encryption';
