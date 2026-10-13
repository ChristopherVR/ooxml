import type { RibbonHomeIntent, ToolbarActionId } from 'ooxml-ui/pptx';
import { pictureSizeHomeControls } from 'ooxml-ui/pptx';
import React, { useCallback, useMemo } from 'react';

import { useToolbarVisibility } from '../../hooks/useToolbarVisibility';
import { useShapeFormatContext } from '../shape-format-context';
import { runCropRequest } from './crop-request';
import { WebHomeControls } from './WebHomeControls';

/**
 * Picture Format > Size: the shared `pptx-ui-ribbon-home-picture-size` strip,
 * whose Crop split button drives the same crop controller as Home > Arrange.
 */
export function PictureSizeStrip({
	hiddenActions,
}: {
	hiddenActions?: readonly ToolbarActionId[];
}): React.ReactElement {
	const crop = useShapeFormatContext()?.crop;
	const { isHidden } = useToolbarVisibility(hiddenActions);
	const controls = useMemo(
		() =>
			pictureSizeHomeControls({
				editable: Boolean(crop),
				canCrop: Boolean(crop?.canCrop),
				cropActive: Boolean(crop?.element),
				hideCrop: isHidden('crop'),
			}),
		[crop, isHidden],
	);
	const request = useCallback(
		(_id: string, _part?: string, intent?: RibbonHomeIntent) => runCropRequest(crop, intent?.value),
		[crop],
	);
	return <WebHomeControls family='picture-size' controls={controls} onRequest={request} />;
}
