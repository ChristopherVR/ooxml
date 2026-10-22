import { defineComponent, h, onBeforeUnmount, onMounted, ref, watch, type PropType } from 'vue';
import type { Workbook } from '@christophervr/xlsx-core';
import type {
	EditorThemeMode,
	RibbonAddInCommandDetail,
	RibbonAddInTab,
	XlsxCollaborationOptions,
	XlsxCollaborationState,
	XlsxEditorElement,
	XlsxThemeColors,
	ExcelCustomization,
} from 'xlsx-web-component';
import {
	EDITOR_EVENT_NAMES,
	EDITOR_PROP_KEYS,
	deferredHandle,
	eventOptions,
	mountEditor,
	pickEditorProps,
	type EditorBinding,
	type EditorPropKey,
} from './index';
export const SpreadsheetEditor = defineComponent({
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
		customization: Object as PropType<ExcelCustomization | null>,
	} satisfies Record<EditorPropKey, unknown>,
	emits: [...EDITOR_EVENT_NAMES, 'update:readOnly', 'ready'],
	setup(props, { emit, expose }) {
		const host = ref<HTMLElement>();
		let binding: EditorBinding | undefined;
		const options = () => ({
			...pickEditorProps(props),
			...eventOptions({
				'workbook-change': (workbook: Workbook) => emit('workbook-change', workbook),
				'workbook-error': (error: Error) => emit('workbook-error', error),
				'selection-change': (detail) => emit('selection-change', detail),
				'dirty-change': (dirty: boolean) => emit('dirty-change', dirty),
				'readonly-change': (readOnly, detail) => {
					emit('readonly-change', readOnly, detail);
					// `v-model:readOnly` mirrors the host's switch only, never a role lock.
					if (detail.hostReadOnly !== props.readOnly) emit('update:readOnly', detail.hostReadOnly);
				},
				'file-command': (event) => emit('file-command', event),
				'workbook-warning': (warnings) => emit('workbook-warning', warnings),
				'ribbon-action': (detail) => emit('ribbon-action', detail),
				'ribbon-customize': (hiddenActions: string[]) => emit('ribbon-customize', hiddenActions),
				'sheet-change': (detail) => emit('sheet-change', detail),
				'office-ribbon-add-in': (detail: RibbonAddInCommandDetail) =>
					emit('office-ribbon-add-in', detail),
				'collaboration-change': (state: XlsxCollaborationState) =>
					emit('collaboration-change', state),
				ready: (element: XlsxEditorElement) => emit('ready', element),
			}),
		});
		onMounted(() => {
			binding = mountEditor(host.value!, options());
		});
		watch(
			() => EDITOR_PROP_KEYS.map((key) => props[key]),
			() => binding?.update(options()),
		);
		onBeforeUnmount(() => binding?.destroy());
		const handle = deferredHandle(() => binding);
		expose({
			get element() {
				return binding?.element;
			},
			load: handle.load,
			newWorkbook: handle.newWorkbook,
			save: handle.save,
			saveBytes: handle.saveBytes,
			download: handle.download,
			markClean: handle.markClean,
			select: handle.select,
			getSelection: handle.getSelection,
			setActiveSheet: handle.setActiveSheet,
			share: handle.share,
			get dirty() {
				return handle.dirty;
			},
		});
		return () => h('div', { ref: host });
	},
});
