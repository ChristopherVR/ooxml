import type { RibbonGroupId } from 'ooxml-ui/pptx';
import { RIBBON_GROUP_ATTR } from 'ooxml-ui/pptx';
import React from 'react';

/** `data-ribbon-group` for an optional catalogue id (nothing when absent). */
export function groupAttr(id: RibbonGroupId | undefined): Record<string, string> {
	return id ? { [RIBBON_GROUP_ATTR]: id } : {};
}

/**
 * A layout-neutral (`display: contents`) wrapper naming a ribbon group, for
 * groups whose markup has no wrapper of its own. Several scopes may carry the
 * same id when a group's controls are not contiguous.
 */
export function RibbonGroupScope({
	id,
	children,
}: {
	id: RibbonGroupId;
	children: React.ReactNode;
}): React.ReactElement {
	return (
		<div className='contents' {...groupAttr(id)}>
			{children}
		</div>
	);
}
