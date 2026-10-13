<script lang="ts">
	/**
	 * PictureSizeStrip: Picture Format > Size, the shared
	 * `pptx-ui-ribbon-home-picture-size` strip. Its Crop split button drives the
	 * editor's crop controller exactly as Home > Arrange's does.
	 */
	import { homeSnapshotTranslator, isActionHidden, pictureSizeHomeControls } from 'ooxml-ui/pptx';
	import type { RibbonHomeRequestEvent, ToolbarActionId } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../../i18n/context';
	import type { EditorState } from '../../../editor/editor-state.svelte';
	import { refocusViewerRoot } from '../anchored-popup';
	import { runCropRequest } from '../home/home-crop';

	const {
		editor,
		hiddenActions,
	}: { editor: EditorState; hiddenActions?: readonly ToolbarActionId[] } = $props();
	const t = useTranslator();
	const state = $derived({
		controls: pictureSizeHomeControls({
			editable: editor.editable,
			canCrop: editor.cropOps.canCrop,
			cropActive: editor.cropOps.active,
			hideCrop: isActionHidden('crop', hiddenActions),
		}),
		translate: homeSnapshotTranslator(['picture-size'], t),
	});

	function request(event: RibbonHomeRequestEvent): void {
		const { value } = event.detail;
		if (value !== undefined) {
			refocusViewerRoot(event.currentTarget as HTMLElement);
		}
		runCropRequest(editor, value);
	}
</script>

<pptx-ui-ribbon-home-picture-size {state} onhome-request={request}></pptx-ui-ribbon-home-picture-size>
