import type { PictureRibbonState, RibbonContextualTabId, ToolbarActionId } from 'ooxml-ui/pptx';
import { CONTEXTUAL_TAB_GROUPS } from 'ooxml-ui/pptx';
import React from 'react';

import { ChartSelectionStrip } from './ChartSelectionStrip';
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
	picture,
	arrange,
}: {
	tab: RibbonContextualTabId;
	hiddenActions?: readonly ToolbarActionId[];
	picture?: PictureRibbonState | undefined;
	arrange?: React.ReactNode;
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
					{group.strip === 'chart-selection' && <ChartSelectionStrip />}
					{(group.pictureStrip === 'picture-arrange' || group.arrange) && arrange}
					{group.contextualStrip &&
						React.createElement('pptx-ui-contextual-ribbon', {
							ref: (node: HTMLElementTagNameMap['pptx-ui-contextual-ribbon'] | null) => {
								if (!node) {
									return;
								}
								node.group = group.group;
								node.state = picture;
							},
						})}
					{group.pictureStrip &&
						React.createElement('pptx-ui-picture-ribbon', {
							ref: (node: HTMLElementTagNameMap['pptx-ui-picture-ribbon'] | null) => {
								if (!node || !group.pictureStrip) {
									return;
								}
								node.family = group.pictureStrip;
								node.state = picture;
							},
						})}
				</WebRibbonGroup>
			))}
		</>
	);
}
