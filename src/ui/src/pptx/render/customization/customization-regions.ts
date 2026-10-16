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
export type HostBackstageRegionId = `backstage.host.${string}`;

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
export interface ViewerRegionMountResult {
	/** Called with the new context whenever it changes. */
	update?(context: ViewerRegionContext): void;
	/** Called once when the region unmounts. */
	destroy?(): void;
}

/** Fill `container` for one region. Return a cleanup function or a handle. */
export type ViewerRegionMount = (
	container: HTMLElement,
	context: ViewerRegionContext,
) => void | (() => void) | ViewerRegionMountResult;

/** A custom element to create in the region (registered by the host). */
export interface ViewerRegionElement {
	/** The tag name, for example `my-settings-view`. */
	element: string;
	/** Properties assigned to the element once, before it is attached. */
	properties?: Readonly<Record<string, unknown>>;
}

/** Any of the renderer forms a region accepts. A string is a tag name. */
export type ViewerRegionRenderer = ViewerRegionMount | ViewerRegionElement | string;

/** A host page added to the File tab, after the built-in pages of its group. */
export interface HostBackstagePage {
	/** Unique among host pages; the region id is `backstage.host.<id>`. */
	id: string;
	/** The navigation label (the host localises it). */
	label: string;
	/** `footer` places it next to Account and Options; default `main`. */
	group?: 'main' | 'footer';
	render: ViewerRegionRenderer;
}

/** Attribute every region container carries, for host CSS and tests. */
export const VIEWER_REGION_ATTRIBUTE = 'data-pptx-region';

/**
 * Mounts one region renderer into one container and keeps it current.
 *
 * A binding creates one host per container (on mount), calls `render` with the
 * current renderer and context on every change, and `destroy` on unmount.
 * `render` remounts only when the renderer itself changes; a new context goes
 * to `update` / the `context` property, and only when a field changed.
 */
export class ViewerRegionHost {
	readonly #container: HTMLElement;
	#renderer: ViewerRegionRenderer | undefined;
	#context: ViewerRegionContext | undefined;
	#result: ViewerRegionMountResult | undefined;
	#element: (HTMLElement & { context?: ViewerRegionContext }) | undefined;
	/** The latest context as the binding passed it, callbacks included. */
	#latest: ViewerRegionContext | undefined;

	constructor(container: HTMLElement) {
		this.#container = container;
	}

	/** Mount `renderer`, or deliver a changed `context` to the mounted one. */
	render(renderer: ViewerRegionRenderer, latest: ViewerRegionContext): void {
		this.#latest = latest;
		const context = this.#stable(latest);
		this.#container.setAttribute(VIEWER_REGION_ATTRIBUTE, context.region);
		if (!sameRenderer(renderer, this.#renderer)) {
			this.#unmount();
			this.#renderer = renderer;
			this.#context = context;
			this.#mount(renderer, context);
			return;
		}
		if (this.#context && sameRegionContext(this.#context, context)) {
			return;
		}
		this.#context = context;
		if (this.#element) {
			this.#element.context = context;
		} else {
			guard(() => this.#result?.update?.(context));
		}
	}

	/** Tear the renderer down and empty the container. */
	destroy(): void {
		this.#unmount();
		this.#renderer = undefined;
		this.#context = undefined;
		this.#latest = undefined;
	}

	/**
	 * `context` with its callbacks swapped for stable ones that call the
	 * latest binding callback, so a renderer that kept an older context never
	 * calls a stale closure (callback identity is not a context change).
	 */
	#stable(context: ViewerRegionContext): ViewerRegionContext {
		const latest = (): Record<string, unknown> =>
			(this.#latest ?? context) as unknown as Record<string, unknown>;
		const call =
			(key: string) =>
			(...args: unknown[]): void =>
				(latest()[key] as ((...a: unknown[]) => void) | undefined)?.(...args);
		const stable: Record<string, unknown> = { ...context, close: call('close') };
		for (const key of CALLBACK_KEYS) {
			if (key !== 'close' && key in context) {
				stable[key] = call(key);
			}
		}
		return stable as unknown as ViewerRegionContext;
	}

	#mount(renderer: ViewerRegionRenderer, context: ViewerRegionContext): void {
		if (typeof renderer === 'function') {
			const result = guard(() => renderer(this.#container, context));
			this.#result = typeof result === 'function' ? { destroy: result } : (result ?? undefined);
			return;
		}
		const spec = typeof renderer === 'string' ? { element: renderer } : renderer;
		const element = guard(() => this.#container.ownerDocument.createElement(spec.element));
		if (!element) {
			return;
		}
		Object.assign(element, spec.properties);
		this.#element = element as HTMLElement & { context?: ViewerRegionContext };
		this.#element.context = context;
		this.#container.append(element);
	}

	#unmount(): void {
		const result = this.#result;
		this.#result = undefined;
		guard(() => result?.destroy?.());
		this.#element?.remove();
		this.#element = undefined;
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
	if (a === b) {
		return true;
	}
	if (!a || !b || typeof a !== 'object' || typeof b !== 'object' || a.element !== b.element) {
		return false;
	}
	const left = a.properties ?? {};
	const right = b.properties ?? {};
	const keys = Object.keys(left);
	return (
		keys.length === Object.keys(right).length &&
		keys.every((key) => Object.is(left[key], right[key]))
	);
}

/** Context fields that are callbacks, never compared (see `#stable`). */
const CALLBACK_KEYS: ReadonlySet<string> = new Set(['close', 'setOption', 'resetOptions']);

/** Shallow equality over the context fields (the selection compared by value). */
export function sameRegionContext(a: ViewerRegionContext, b: ViewerRegionContext): boolean {
	const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
	for (const key of keys) {
		const left = (a as unknown as Record<string, unknown>)[key];
		const right = (b as unknown as Record<string, unknown>)[key];
		if (key === 'selectedElementIds') {
			if (!sameIds(left as readonly string[], right as readonly string[])) {
				return false;
			}
		} else if (CALLBACK_KEYS.has(key)) {
			// Callbacks are recreated on every render; their identity is not a change.
			continue;
		} else if (!Object.is(left, right)) {
			return false;
		}
	}
	return true;
}

function sameIds(a: readonly string[] | undefined, b: readonly string[] | undefined): boolean {
	if (a === b) {
		return true;
	}
	if (!a || !b || a.length !== b.length) {
		return false;
	}
	return a.every((id, index) => id === b[index]);
}

/** A host renderer must never take the viewer down with it. */
function guard<T>(run: () => T): T | undefined {
	try {
		return run();
	} catch (error) {
		console.error('[pptx-viewer] A host region renderer threw:', error);
		return undefined;
	}
}
