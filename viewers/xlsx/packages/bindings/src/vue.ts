import type { PropType } from 'vue';
import type { Workbook } from '@christophervr/xlsx-core';
import type { DocumentStorage } from 'ooxml-core/embedding';
import type {
	EditorThemeMode,
	RibbonAddInTab,
	XlsxCollaborationOptions,
	XlsxThemeColors,
	ExcelCustomization,
} from 'xlsx-web-component';
import { defineVueBinding } from 'ooxml-ui/bindings/vue';
import { EDITOR_CONTRACT, mountEditor, type EditorPropKey } from './index';
/**
 * Kebab-case emits for every editor event plus `ready`, `v-model:readOnly` (the host's switch,
 * never a role lock) and the handle on the template ref, all generated from `EDITOR_CONTRACT`.
 */
export const SpreadsheetEditor = defineVueBinding({
	name: 'SpreadsheetEditor',
	props: {
		workbook: Object as PropType<Workbook>,
		bytes: Object as PropType<Uint8Array | ArrayBuffer>,
		src: String,
		fileName: String,
		readOnly: Boolean,
		locale: String,
		theme: String as PropType<EditorThemeMode>,
		authorName: String,
		showToolbar: { type: Boolean, default: true },
		showFormulaBar: { type: Boolean, default: true },
		hiddenActions: Array as PropType<readonly string[]>,
		ribbonAddIns: Array as unknown as PropType<readonly RibbonAddInTab[]>,
		themeColors: Object as PropType<XlsxThemeColors>,
		collaboration: Object as PropType<XlsxCollaborationOptions | null>,
		storage: Object as PropType<DocumentStorage | null>,
		customization: Object as PropType<ExcelCustomization | null>,
	} satisfies Record<EditorPropKey, unknown>,
	contract: EDITOR_CONTRACT,
	callbacks: { ready: 'onReady' },
	handle: EDITOR_CONTRACT.handle,
	mount: mountEditor,
});
