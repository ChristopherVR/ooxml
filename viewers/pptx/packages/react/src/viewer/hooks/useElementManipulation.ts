/**
 * useElementManipulation: Clipboard (copy/cut/paste/duplicate), group/ungroup,
 * delete, flip, align, layer-order, and context-menu dispatch handlers.
 *
 * Handler logic is split across sub-hooks; this module composes them and
 * provides the context-menu dispatch.
 */
import { contextMenuInspectorAnchor, formatSectionForAnchor } from 'ooxml-ui/pptx';
import type { FormatPaneController } from 'ooxml-ui/pptx';

import type { ElementContextMenuAction } from '../types';
import type {
	UseElementManipulationInput,
	ElementManipulationHandlers,
} from './element-manipulation-types';
import { useClipboardHandlers } from './useClipboardHandlers';
import { useGroupAlignLayerHandlers } from './useGroupAlignLayerHandlers';

/** Open the Format pane on the section `anchor` names (the request opens the tab). */
function focusInspectorSection(
	formatPane: FormatPaneController,
	anchor: ReturnType<typeof contextMenuInspectorAnchor>,
): void {
	formatPane.requestFormatPane(anchor ? formatSectionForAnchor(anchor) : null, { focus: true });
}

export type {
	UseElementManipulationInput,
	ElementManipulationHandlers,
} from './element-manipulation-types';

export function useElementManipulation(
	input: UseElementManipulationInput,
): ElementManipulationHandlers {
	const {
		selectedElement,
		setIsInspectorPaneOpen,
		setSidebarPanelMode,
		formatPane,
		onOpenHyperlinkDialog,
		onEditText,
		onEditPoints,
		onSaveElementAsPicture,
	} = input;

	const { handleCopy, handleCut, handlePaste, handleDuplicate, handleDelete } =
		useClipboardHandlers(input);

	const {
		handleGroupElements,
		handleUngroupElement,
		handleFlip,
		handleAlignElements,
		handleDistributeElements,
		canDistribute,
		selectionGroupable,
		handleMoveLayer,
		handleMoveLayerToEdge,
		handleMergeShapes,
		canMergeShapes,
	} = useGroupAlignLayerHandlers(input);

	const handleContextMenuAction = (action: ElementContextMenuAction) => {
		switch (action) {
			case 'copy':
				handleCopy();
				break;
			case 'cut':
				handleCut();
				break;
			case 'paste':
				handlePaste();
				break;
			case 'duplicate':
				handleDuplicate();
				break;
			case 'delete':
				handleDelete();
				break;
			case 'bring-forward':
			case 'bringForward':
				handleMoveLayer('forward');
				break;
			case 'send-backward':
			case 'sendBackward':
				handleMoveLayer('backward');
				break;
			case 'bring-front':
			case 'bringToFront':
				handleMoveLayerToEdge('front');
				break;
			case 'send-back':
			case 'sendToBack':
				handleMoveLayerToEdge('back');
				break;
			case 'comment':
			case 'addComment':
				setIsInspectorPaneOpen(true);
				setSidebarPanelMode('comments');
				break;
			case 'group':
				handleGroupElements();
				break;
			case 'ungroup':
				handleUngroupElement();
				break;
			case 'editHyperlink':
				onOpenHyperlinkDialog();
				break;
			case 'edit-text':
				if (selectedElement) {
					onEditText?.(selectedElement.id);
				}
				break;
			case 'edit-points':
			case 'editPoints':
				if (selectedElement) {
					onEditPoints?.(selectedElement.id);
				}
				break;
			case 'save-as-picture':
				if (selectedElement) {
					onSaveElementAsPicture?.(selectedElement.id);
				}
				break;
			case 'edit-alt-text':
				focusInspectorSection(formatPane, contextMenuInspectorAnchor('edit-alt-text'));
				break;
			case 'size-and-position':
				focusInspectorSection(formatPane, contextMenuInspectorAnchor('size-and-position'));
				break;
			case 'format-shape':
				focusInspectorSection(formatPane, contextMenuInspectorAnchor('format-shape'));
				break;
		}
	};

	return {
		handleCopy,
		handleCut,
		handlePaste,
		handleDuplicate,
		handleGroupElements,
		handleUngroupElement,
		handleDelete,
		handleFlip,
		handleAlignElements,
		handleDistributeElements,
		canDistribute,
		selectionGroupable,
		handleMoveLayer,
		handleMoveLayerToEdge,
		handleMergeShapes,
		canMergeShapes,
		handleContextMenuAction,
	};
}
