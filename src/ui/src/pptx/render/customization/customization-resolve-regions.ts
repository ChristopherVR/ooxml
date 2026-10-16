/**
 * Which replaced regions survive resolution: a region the host also hid
 * (as a panel, a dialog or a File page) is dropped, so hiding always wins.
 *
 * @module render/customization/customization-resolve-regions
 */
import type { BackstagePage } from '../backstage';
import type { ViewerRegionId, ViewerRegionRenderer } from './customization-regions';
import type { ViewerCustomization, ViewerDialogId, ViewerPanelId } from './customization-types';

/** The region a dialog id stands for, when it is replaceable. */
const REGION_OF_DIALOG: Partial<Record<ViewerDialogId, ViewerRegionId>> = {
	options: 'dialog.options',
	share: 'dialog.share',
	broadcast: 'dialog.broadcast',
	print: 'dialog.print',
};

export function resolveRegions(
	c: ViewerCustomization,
	hiddenPanels: ReadonlySet<ViewerPanelId>,
	hiddenDialogs: ReadonlySet<ViewerDialogId>,
	hiddenPages: ReadonlySet<BackstagePage>,
): Map<ViewerRegionId, ViewerRegionRenderer> {
	const hidden = new Set<string>(hiddenPanels);
	for (const dialog of hiddenDialogs) {
		const region = REGION_OF_DIALOG[dialog];
		if (region) {
			hidden.add(region);
		}
	}
	for (const page of hiddenPages) {
		hidden.add(`backstage.${page}`);
	}
	const regions = new Map<ViewerRegionId, ViewerRegionRenderer>();
	for (const [id, renderer] of Object.entries(c.regions ?? {}) as Array<
		[ViewerRegionId, ViewerRegionRenderer | undefined]
	>) {
		if (renderer && !hidden.has(id)) {
			regions.set(id, renderer);
		}
	}
	return regions;
}
