import type { RibbonContextualTabId, ToolbarActionId } from 'ooxml-ui/pptx';
import { CONTEXTUAL_TAB_GROUPS } from 'ooxml-ui/pptx';
import React from 'react';

import { PictureSizeStrip } from './PictureSizeStrip';
import { RibbonGallery } from './RibbonGallery';
import { useTranslateOr } from './useGalleryTranslation';
import { WebRibbonGroup } from './WebRibbonControls';

/**
 * A contextual tab's ribbon content (Shape Format, Picture Format, Table
 * Design, ...): the groups shared lists for it, each holding its galleries
 * and, for a command group such as Picture Format > Size, its shared strip.
 */
export function ContextualTabSection({
	tab,
	hiddenActions,
}: {
	tab: RibbonContextualTabId;
	hiddenActions?: readonly ToolbarActionId[];
}): React.ReactElement {
	const translateOr = useTranslateOr();
	return (
		<>
			{CONTEXTUAL_TAB_GROUPS[tab].map((group) => (
				<WebRibbonGroup
					key={group.group}
					groupId={group.group}
					label={translateOr(group.labelKey, group.label)}
				>
					{group.galleries.map((placement) => (
						<RibbonGallery key={placement.control} placement={placement} />
					))}
					{group.strip === 'picture-size' && <PictureSizeStrip hiddenActions={hiddenActions} />}
				</WebRibbonGroup>
			))}
		</>
	);
}
