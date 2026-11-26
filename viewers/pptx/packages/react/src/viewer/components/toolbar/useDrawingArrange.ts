import type { DrawingArrangeInput } from 'ooxml-ui/pptx';
import { useMemo } from 'react';

import type { useFormatPane } from '../inspector/format-pane-context';
import type { ToolbarProps } from './toolbar-types';

export function useDrawingArrange(
	props: ToolbarProps,
	formatPane: ReturnType<typeof useFormatPane>,
): DrawingArrangeInput {
	const {
		canEdit,
		onAlignElements,
		onDistributeElements,
		onFlip,
		onGroupElements,
		onPatchElement,
		onToggleSelectionPane,
		onUngroupElement,
		selectedCount,
		selectedElement,
		selectionGroupable,
	} = props;
	return useMemo<DrawingArrangeInput>(
		() => ({
			actions: {
				align: (edge) => onAlignElements(edge === 'centerH' ? 'center' : edge),
				distribute: onDistributeElements,
				flip: onFlip,
				rotation: onPatchElement
					? () => ({ element: selectedElement, patch: onPatchElement })
					: undefined,
				group: onGroupElements,
				ungroup: onUngroupElement,
				position: () => formatPane.requestFormatPane('sizePosition', { focus: true }),
				selectionPane: onToggleSelectionPane,
			},
			selectedCount,
			selectedElement,
			canEdit,
			canGroup: selectedCount > 1 && selectionGroupable,
		}),
		[
			formatPane,
			onPatchElement,
			canEdit,
			onAlignElements,
			onDistributeElements,
			onFlip,
			onGroupElements,
			onToggleSelectionPane,
			onUngroupElement,
			selectedCount,
			selectedElement,
			selectionGroupable,
		],
	);
}
