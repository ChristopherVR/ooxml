// Compatibility exports: document operations live in ooxml-core.
import { saveDeckWithPassword as saveDeckThroughPlan } from 'ooxml-core/pptx/editor/render/deck-save-encryption';
import type {
	DeckSaveIntent,
	DeckSaveOptions,
	DeckSaveSerializer,
} from 'ooxml-core/pptx/editor/render/deck-save-encryption';
import type { PptxSlide } from 'ooxml-core/pptx';

import { settleThemeWrites } from './theme-history';

export {
	planDeckSave,
	recoverySnapshotIntent,
	isEncryptedDeckBytes,
} from 'ooxml-core/pptx/editor/render/deck-save-encryption';
export type {
	DeckSaveSerializer,
	DeckSaveOptions,
	DeckSavePurpose,
	DeckSaveIntent,
	DeckSavePlan,
} from 'ooxml-core/pptx/editor/render/deck-save-encryption';

/**
 * The core's `saveDeckWithPassword` (the shared `planDeckSave` decision: a captured password means
 * `saveEncrypted`, never `save`), once every theme part write queued on `serializer` has landed.
 * Every binding's save, export, autosave and recovery snapshot comes through here, so none of them
 * can write a theme an undo or a Design gallery pick is still replacing (`render/theme-history`).
 */
export async function saveDeckWithPassword(
	serializer: DeckSaveSerializer,
	slides: PptxSlide[],
	options: DeckSaveOptions | undefined,
	intent: DeckSaveIntent | string | null | undefined,
): Promise<Uint8Array> {
	await settleThemeWrites(serializer);
	return saveDeckThroughPlan(serializer, slides, options, intent);
}
