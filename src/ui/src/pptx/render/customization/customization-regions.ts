/**
 * Host-supplied UI: the "replace" half of the customisation model.
 *
 * Hiding a region (`hiddenPanels`, `hiddenDialogs`, ...) removes it. Replacing
 * a region keeps its place in the layout (the inspector's column, the Options
 * dialog's frame, a File tab page's body) and hands that place to the host as
 * an empty container element. The host fills it with whatever it likes, in
 * any framework, through ONE of two renderer forms:
 *
 *  - a mount function `(container, context) => cleanup | handle`, or
 *  - a custom element: its tag name, or `{ element, properties }`. The viewer
 *    creates it, assigns `properties`, and keeps its `context` property current.
 *
 * Every binding hosts a region through `ViewerRegionHost` below, so the mount,
 * update and teardown semantics are identical in React, Vue, Angular, Svelte
 * and Vanilla; a binding only decides WHERE the container goes.
 *
 * @module render/customization/customization-regions
 */
import type { PowerPointViewerAPI, ViewerMode } from '../../types';
import type { BackstagePage } from '../backstage';
import type { ViewerOptionPrimitive, ViewerOptions } from '../options/viewer-options';
import type { OptionsSettingId } from './customization-types';
import {
	sameRegionContext as sameHostRegionContext,
	sameRenderer as sameHostRenderer,
	type HostBackstagePage as CoreHostBackstagePage,
	type HostBackstageRegionId,
	type HostRegionElement,
	type HostRegionMount,
	type HostRegionMountResult,
	type HostRegionRenderer,
} from 'ooxml-core/embedding';
import { HostRegionHost } from '../../../embedding/region-host';

/** Docked panels a host can replace (the whole panel, keeping its slot). */
export type ViewerPanelRegionId = 'inspector' | 'slidesPane' | 'notes' | 'statusBar';

/**
 * Dialogs a host can replace. The viewer keeps the dialog frame (backdrop,
 * title bar, close button, Escape) and replaces its body. `dialog.options` is
 * File > Options, the Settings dialog.
 */
export type ViewerDialogRegionId =
	| 'dialog.options'
	| 'dialog.share'
	| 'dialog.broadcast'
	| 'dialog.print';

/** File tab (backstage) pages a host can replace: the page body, not its nav entry. */
export type ViewerBackstageRegionId = `backstage.${BackstagePage}`;

/** Every built-in region a host can replace. */
export type ViewerRegionId = ViewerPanelRegionId | ViewerDialogRegionId | ViewerBackstageRegionId;

/** The region id a host File tab page (`backstage.extraPages`) is mounted under. */
export type { HostBackstageRegionId };

/**
 * What a host renderer receives. It is rebuilt whenever any field changes and
 * delivered through `update` (mount functions) or the `context` property
 * (custom elements).
 */
export interface ViewerRegionContext {
	/** Which region this renderer fills. */
	readonly region: ViewerRegionId | HostBackstageRegionId;
	/**
	 * The viewer's imperative API (the same object the component handle
	 * exposes): read slides and elements, update them, navigate, undo, ...
	 */
	readonly viewer: PowerPointViewerAPI;
	readonly mode: ViewerMode;
	/** False when the deck is read-only (prop, capability or protection). */
	readonly canEdit: boolean;
	readonly activeSlideIndex: number;
	readonly slideCount: number;
	/** Ids of the selected elements on the active slide (empty: no selection). */
	readonly selectedElementIds: readonly string[];
	/** The active UI locale (BCP 47). */
	readonly locale: string;
	/**
	 * Dismiss the region: closes a dialog, leaves the File tab, or closes the
	 * inspector. A no-op for regions that cannot close (status bar, notes,
	 * slides pane).
	 */
	readonly close: () => void;
}

/** The context of `dialog.options`: the File > Options model to edit. */
export interface ViewerSettingsRegionContext extends ViewerRegionContext {
	readonly region: 'dialog.options';
	/** The user's current File > Options values (host locks applied). */
	readonly options: ViewerOptions;
	/** Change one setting; locked settings ignore this, as in the built-in dialog. */
	readonly setOption: (setting: OptionsSettingId, value: ViewerOptionPrimitive) => void;
	/** Reset every setting to the host defaults, or the built-ins. */
	readonly resetOptions: () => void;
}

/** What a mount function may return to receive updates and teardown. */
export type ViewerRegionMountResult = HostRegionMountResult<ViewerRegionContext>;

/** Fill `container` for one region. Return a cleanup function or a handle. */
export type ViewerRegionMount = HostRegionMount<HTMLElement, ViewerRegionContext>;

/** A custom element to create in the region (registered by the host). */
export type ViewerRegionElement = HostRegionElement;

/** Any of the renderer forms a region accepts. A string is a tag name. */
export type ViewerRegionRenderer = HostRegionRenderer<HTMLElement, ViewerRegionContext>;

/**
 * A host page added to the File tab, after the built-in pages of its group (the shared
 * `HostBackstagePage` of `ooxml-core/embedding`, with PowerPoint's region renderer).
 */
export type HostBackstagePage = CoreHostBackstagePage<ViewerRegionRenderer>;

/** Attribute every region container carries, for host CSS and tests. */
export const VIEWER_REGION_ATTRIBUTE = 'data-pptx-region';

/** Context fields that are callbacks, never compared (`close` is always one). */
const CALLBACK_KEYS: ReadonlySet<string> = new Set(['close', 'setOption', 'resetOptions']);

/**
 * Mounts one region renderer into one container and keeps it current: the shared
 * `HostRegionHost` of `ooxml-ui` with PowerPoint's region attribute and settings callbacks.
 *
 * A binding creates one host per container (on mount), calls `render` with the
 * current renderer and context on every change, and `destroy` on unmount.
 * `render` remounts only when the renderer itself changes; a new context goes
 * to `update` / the `context` property, and only when a field changed.
 */
export class ViewerRegionHost extends HostRegionHost<ViewerRegionContext> {
	constructor(container: HTMLElement) {
		super(container, {
			attribute: VIEWER_REGION_ATTRIBUTE,
			callbackKeys: CALLBACK_KEYS,
			label: '[pptx-viewer]',
		});
	}
}

/**
 * True when two renderers are the same: the same function, the same tag, or
 * element specs with the same tag and the same property values (so a host can
 * pass an inline `{ element, properties }` without remounting every render).
 */
export function sameRenderer(
	a: ViewerRegionRenderer | undefined,
	b: ViewerRegionRenderer | undefined,
): boolean {
	return sameHostRenderer(a, b);
}

/** Shallow equality over the context fields (the selection compared by value). */
export function sameRegionContext(a: ViewerRegionContext, b: ViewerRegionContext): boolean {
	return sameHostRegionContext(a, b, CALLBACK_KEYS);
}
