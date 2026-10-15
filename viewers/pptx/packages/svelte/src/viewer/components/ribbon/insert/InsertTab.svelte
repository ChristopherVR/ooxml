<script lang="ts">
	/**
	 * InsertTab: thin adapter for the shared `pptx-ui-ribbon-insert`. The shared
	 * element owns the groups, icons, labels, shape/chart pickers, Freeform tools and
	 * the Action / Field menus; this component supplies viewer state and routes typed
	 * intents to the native handlers. Every insertion routes through
	 * `EditorState.insertElement` (undoable, selects the new element). The file
	 * pickers, the SmartArt gallery, the equation editor and the hyperlink dialog stay
	 * native to the Svelte binding.
	 */
	import type { PptxLayoutOption, SmartArtLayout } from 'pptx-viewer-core';
	import {
		DEFAULT_INSERT_CHART_KIND,
		FREEFORM_TOOL_IDS,
		isDrawingToolVisible,
		resolveInsertSlideLayout,
		scopeLayoutOptionsToSlide,
	} from 'ooxml-ui/pptx';
	import type {
		CanvasSize,
		FreeformToolKind,
		InsertChartKind,
		PptxUiRibbonInsertElement,
		RibbonInsertRequestEvent,
		ShapePresetType,
	} from 'ooxml-ui/pptx';
	import { tick } from 'svelte';

	import { useTranslator } from '../../../../i18n/context';
	import MobileSheet from '../../MobileSheet.svelte';
	import {
		buildActionButtonInsertElement,
		buildChartInsertElement,
		buildFieldInsertElement,
		buildMediaInsertElement,
		buildSmartArtInsertElement,
		newImageElement,
		newPresetShapeElement,
		newTableElement,
		newTextElement,
		resolveFieldDisplayText,
	} from '../../../editor';
	import type { EditorState } from '../../../editor/editor-state.svelte';
	import type { ChromeUiState } from '../../../state/chrome-ui.svelte';
	import { useViewerCustomization } from '../../../state/viewer-customization.svelte';
	import EquationEditorDialog from './EquationEditorDialog.svelte';
	import HyperlinkDialog from './HyperlinkDialog.svelte';
	import SmartArtDialog from './SmartArtDialog.svelte';

	const {
		editor,
		canvasSize,
		onheaderfooter,
		onnavigateslide,
		chromeUi,
		phone = false,
		open = true,
		onclose,
	}: {
		editor: EditorState;
		canvasSize: CanvasSize;
		onheaderfooter?: () => void;
		/** Select the slide New Slide just added. */
		onnavigateslide?: (index: number) => void;
		/** Comment opens the inspector's Comments pane, as Review > New Comment does. */
		chromeUi?: ChromeUiState;
		/**
		 * Draw the phone Insert sheet (`pptx-ui-mobile-insert` in a bottom sheet) instead of the ribbon
		 * group row. The dialogs below stay mounted while the sheet is closed, so a pick that opens
		 * one (SmartArt, equation, link) can close the sheet first.
		 */
		phone?: boolean;
		open?: boolean;
		onclose?: () => void;
	} = $props();
	const t = useTranslator();
	const customization = useViewerCustomization();

	// eslint-disable-next-line prefer-const
	let host = $state<PptxUiRibbonInsertElement | null>(null);
	// eslint-disable-next-line prefer-const
	let imageInput = $state<HTMLInputElement | null>(null);
	// eslint-disable-next-line prefer-const
	let mediaInput = $state<HTMLInputElement | null>(null);
	let equationOpen = $state(false);
	let smartArtOpen = $state(false);
	// eslint-disable-next-line prefer-const
	let hyperlinkOpen = $state(false);
	let shapeType = $state<ShapePresetType>('rect');
	let chartKind = $state<InsertChartKind>(DEFAULT_INSERT_CHART_KIND);
	$effect(() => {
		if (editor.equationOps.editingId) {
			equationOpen = true;
		}
	});

	/** The deck's layouts, loaded once so the New Slide chevron can list them. */
	let layouts = $state<PptxLayoutOption[]>([]);
	const currentLayoutPath = $derived(editor.slides[editor.currentSlideIndex]?.layoutPath);
	/** Scoped to the active slide's own master; dedupes same-named layouts across masters. */
	const scopedLayouts = $derived(scopeLayoutOptionsToSlide(layouts, currentLayoutPath));
	$effect(() => {
		void editor.slidesOps.availableLayouts().then((loaded) => (layouts = loaded));
	});

	const insertState = $derived({
		editable: editor.editable,
		hasSelection: Boolean(editor.selectedElementId),
		shapeType,
		chartKind,
		activeFreeformTool: editor.outlineOps.freeformTool,
		freeformTools: FREEFORM_TOOL_IDS.filter((tool) => isDrawingToolVisible(customization.resolved, tool)),
		headerFooterAvailable: Boolean(onheaderfooter),
		// Like Home, New Slide works without layouts (it adds a blank slide).
		newSlideAvailable: true,
		commentAvailable: Boolean(chromeUi),
		layouts: scopedLayouts.map(({ path, name }) => ({ path, name })),
		...(currentLayoutPath ? { currentLayoutPath } : {}),
		translate: t,
	});

	const MAX_IMAGE_EDGE = 400;

	function closeEquationDialog(): void {
		equationOpen = false;
		editor.equationOps.close();
	}

	function closeSmartArt(): void {
		smartArtOpen = false;
		void tick().then(() => host?.focusControl('insert.illustrations.smartArt'));
	}

	function insertSmartArt(layout: SmartArtLayout, defaultItems: string[]): void {
		editor.insertElement(buildSmartArtInsertElement(layout, defaultItems, canvasSize));
		closeSmartArt();
	}

	function insertField(fieldType: string): void {
		const displayText = resolveFieldDisplayText(fieldType, {
			slideNumber: editor.currentSlideIndex + 1,
		});
		editor.insertElement(buildFieldInsertElement(fieldType, displayText, canvasSize));
	}

	function onImageFileChange(event: Event): void {
		const input = event.currentTarget as HTMLInputElement;
		const file = input.files?.[0];
		input.value = '';
		if (!file) {
			return;
		}
		const reader = new FileReader();
		reader.onload = () => {
			const dataUrl = typeof reader.result === 'string' ? reader.result : '';
			if (!dataUrl) {
				return;
			}
			const probe = new Image();
			probe.onload = () => {
				const ratio = Math.min(
					1,
					MAX_IMAGE_EDGE / Math.max(probe.naturalWidth || 1, probe.naturalHeight || 1),
				);
				const w = Math.max(1, Math.round((probe.naturalWidth || MAX_IMAGE_EDGE) * ratio));
				const h = Math.max(1, Math.round((probe.naturalHeight || MAX_IMAGE_EDGE) * ratio));
				editor.insertElement(newImageElement(dataUrl, 120, 120, w, h));
			};
			probe.onerror = () => {
				editor.insertElement(newImageElement(dataUrl, 120, 120, 300, 200));
			};
			probe.src = dataUrl;
		};
		reader.readAsDataURL(file);
	}

	async function onMediaFileChange(event: Event): Promise<void> {
		const input = event.currentTarget as HTMLInputElement;
		const file = input.files?.[0];
		input.value = '';
		if (!file) {
			return;
		}
		const el = await buildMediaInsertElement(file, canvasSize);
		if (el) {
			editor.insertElement(el);
		}
	}

	function request(event: RibbonInsertRequestEvent): void {
		const intent = event.detail;
		switch (intent.kind) {
			case 'command':
				switch (intent.value) {
					case 'textBox': editor.insertElement(newTextElement()); break;
					case 'table': editor.insertElement(newTableElement()); break;
					case 'image': imageInput?.click(); break;
					case 'media': mediaInput?.click(); break;
					case 'smartArt': smartArtOpen = true; break;
					case 'equation': equationOpen = !equationOpen; break;
					case 'link': hyperlinkOpen = true; break;
					case 'headerFooter': onheaderfooter?.(); break;
					case 'comment':
						// Opening an already open pane would toggle it shut.
						if (chromeUi && !chromeUi.commentsOpen) {
							chromeUi.toggleComments();
						}
						break;
				}
				break;
			case 'table':
				editor.insertElement(newTableElement(intent.rows, intent.columns));
				break;
			case 'newSlide': {
				const layout = resolveInsertSlideLayout(scopedLayouts, intent.value, currentLayoutPath);
				const index = layout
					? editor.slidesOps.insertSlideFromLayout(layout.path, layout.name)
					: editor.slidesOps.insertSlideAfterCurrent();
				if (index !== null) {
					onnavigateslide?.(index);
				}
				break;
			}
			case 'shapeType': shapeType = intent.value as ShapePresetType; break;
			case 'shape': editor.insertElement(newPresetShapeElement(intent.value as ShapePresetType)); break;
			case 'chartType': chartKind = intent.value as InsertChartKind; break;
			case 'chart': editor.insertElement(buildChartInsertElement(intent.value as InsertChartKind, canvasSize)); break;
			case 'freeform': editor.outlineOps.armFreeformTool(intent.value as FreeformToolKind | null); break;
			case 'actionButton': {
				const el = buildActionButtonInsertElement(intent.value, canvasSize);
				if (el) {
					editor.insertElement(el);
				}
				break;
			}
			case 'field': insertField(intent.value); break;
		}
		if (phone) {
			onclose?.();
		}
	}
</script>

{#if !phone}
	<pptx-ui-ribbon-insert bind:this={host} state={insertState} oninsert-request={request}></pptx-ui-ribbon-insert>
{:else if open}
	<MobileSheet title={t('pptx.mobileBar.insert')} onclose={() => onclose?.()}>
		<pptx-ui-mobile-insert state={insertState} oninsert-request={request}></pptx-ui-mobile-insert>
	</MobileSheet>
{/if}
<input bind:this={imageInput} type="file" accept="image/*" class="pptx-svelte-inserttab-file" onchange={onImageFileChange} />
<input bind:this={mediaInput} type="file" accept="video/*,audio/*" class="pptx-svelte-inserttab-file" onchange={onMediaFileChange} />

{#if smartArtOpen}<SmartArtDialog oncancel={closeSmartArt} oninsert={insertSmartArt} />{/if}
<EquationEditorDialog {editor} {canvasSize} open={equationOpen && customization.isDialogAvailable('equation')} onclose={closeEquationDialog} />
{#if hyperlinkOpen && customization.isDialogAvailable('hyperlink')}<HyperlinkDialog {editor} onclose={() => (hyperlinkOpen = false)} />{/if}

<style>
	.pptx-svelte-inserttab-file {
		display: none;
	}
</style>
