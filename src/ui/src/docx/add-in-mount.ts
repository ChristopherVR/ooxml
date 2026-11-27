import { watchOfficeAddIns } from '../ribbon/add-in-dialog';
import { syncInstalledRibbonTabs } from '../ribbon/add-in-tabs';
import { loadWordAddInSelection } from './add-in-selection-loader';
import type { EditorCore } from './editor-core';
import { translate, type LocalizationKey } from './localization';

/** Restores the editor's mounted add-ins and keeps them in step with its preferences. */
export function watchWordAddIns(root: ShadowRoot, core: EditorCore): () => void {
	return watchOfficeAddIns(
		root,
		(label) => translate(core.locale, label as LocalizationKey),
		'word',
		() => core.locale,
		() => loadWordAddInSelection(() => core.targetView()),
		syncInstalledRibbonTabs,
	);
}
