import type { PropType } from 'vue';
import type { DocumentModel } from 'docx-core';
import type { DocumentStorage } from 'ooxml-core/embedding';
import type {
	DocxCollaborationOptions,
	EditorTheme,
	EditorThemeMode,
	RibbonActionInput,
	RibbonAddInTab,
	WordCustomization,
} from 'docx-web-component';
import { defineVueBinding } from 'ooxml-ui/bindings/vue';
import { EDITOR_CONTRACT, mountEditor, type EditorPropKey } from './index';
/**
 * Kebab-case emits for every editor event, `v-model:readOnly` (the host's switch, never a role
 * lock) and the handle on the template ref, all generated from `EDITOR_CONTRACT`.
 */
export const WordEditor = defineVueBinding({
	name: 'WordEditor',
	props: {
		documentModel: Object as PropType<DocumentModel>,
		readOnly: Boolean,
		fileName: String,
		reviewAuthor: String,
		themeColors: Object as PropType<Partial<EditorTheme>>,
		locale: String,
		theme: String as PropType<EditorThemeMode>,
		showThumbnails: Boolean,
		showToolbar: { type: Boolean, default: true },
		hiddenActions: Array as PropType<readonly RibbonActionInput[]>,
		ribbonAddIns: Array as unknown as PropType<readonly RibbonAddInTab[]>,
		collaboration: Object as PropType<DocxCollaborationOptions | null>,
		storage: Object as PropType<DocumentStorage | null>,
		customization: Object as PropType<WordCustomization | null>,
	} satisfies Record<EditorPropKey, unknown>,
	contract: EDITOR_CONTRACT,
	handle: EDITOR_CONTRACT.handle,
	mount: mountEditor,
});
