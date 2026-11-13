import { applyBuildingBlockLiveOptions } from 'ooxml-ui/pptx';

import type { PptxViewerSource } from '../viewer/load';
import type { PptxViewerOptions } from '../viewer/types';
import type { BuildingBlocksViewer } from './blocks-viewer';
import type { ViewerBuildingBlocksOptions } from './types';

/** The regions the canvas block leaves out: the host composes its own shell around it. */
const BLOCK_HIDDEN_PANELS = ['statusBar', 'inspector', 'notes', 'titleBar', 'quickAccessToolbar'];

/** What the blocks hear from the viewer besides the host (the shell state follows the session). */
export interface ViewerOptionHooks {
	onCollaborationChange(): void;
}

/** The full viewer's options for building-block options (chrome off, permission from `canEdit`). */
export function viewerOptions(
	options: ViewerBuildingBlocksOptions,
	hooks?: ViewerOptionHooks,
): PptxViewerOptions {
	// The building-block-only options never reach the viewer.
	const {
		content,
		canEdit,
		onRequest: _onRequest,
		fileLimits: _fileLimits,
		onOpenLayoutGallery: _onOpenLayoutGallery,
		...rest
	} = options;
	const hiddenPanels = [...(rest.customization?.hiddenPanels ?? []), ...BLOCK_HIDDEN_PANELS];
	return {
		...rest,
		...(content ? { source: content } : {}),
		editable: canEdit === true,
		showToolbar: false,
		showThumbnails: false,
		showCompatibilityToasts: false,
		customization: { ...rest.customization, hiddenPanels } as PptxViewerOptions['customization'],
		...(hooks
			? {
					onCollaborationChange: (state) => {
						hooks.onCollaborationChange();
						rest.onCollaborationChange?.(state);
					},
					onCollaborationStatus: (status) => {
						hooks.onCollaborationChange();
						rest.onCollaborationStatus?.(status);
					},
				}
			: {}),
	};
}

/**
 * The viewer's options as a live view of the latest building-block options: everything the
 * viewer reads when it uses it (callbacks, `listConversionScope`, `limits` on the next load, ...)
 * follows `update()` without a remount. Mapped once per options object.
 */
export function liveViewerOptions(
	read: () => ViewerBuildingBlocksOptions,
	hooks?: ViewerOptionHooks,
): PptxViewerOptions {
	let source: ViewerBuildingBlocksOptions | undefined;
	let mapped: PptxViewerOptions = {};
	const current = (): PptxViewerOptions => {
		const next = read();
		if (next !== source) {
			source = next;
			mapped = viewerOptions(next, hooks);
		}
		return mapped;
	};
	return new Proxy({} as PptxViewerOptions, {
		get: (_target, key) => Reflect.get(current(), key),
		has: (_target, key) => key in current(),
		ownKeys: () => Reflect.ownKeys(current()),
		getOwnPropertyDescriptor: (_target, key) => {
			const descriptor = Reflect.getOwnPropertyDescriptor(current(), key);
			return descriptor ? { ...descriptor, configurable: true } : undefined;
		},
	});
}

/** The options `update()` has to push to the viewer explicitly (the rest are read live). */
export function applyOptionChanges(
	viewer: BuildingBlocksViewer,
	previous: ViewerBuildingBlocksOptions,
	next: ViewerBuildingBlocksOptions,
): void {
	if ((next.canEdit === true) !== (previous.canEdit === true)) {
		viewer.setEditable(next.canEdit === true);
	}
	if (next.locale !== previous.locale && next.locale) {
		viewer.setLocale(next.locale);
	}
	if (next.theme !== previous.theme) {
		viewer.setTheme(next.theme);
	}
	if (next.customization !== previous.customization) {
		const { customization } = viewerOptions(next);
		if (customization) {
			viewer.setCustomization(customization);
		}
	}
	// The options the vanilla viewer reads once: pushed to the mounted viewer, the deck stays.
	applyBuildingBlockLiveOptions(previous, next, viewer.liveOptionPort);
}

export async function loadSource(
	viewer: BuildingBlocksViewer,
	source: PptxViewerSource,
): Promise<void> {
	await (typeof source === 'string' ? viewer.loadUrl(source) : viewer.loadFile(source));
}
