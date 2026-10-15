/**
 * The vanilla region host: mounts a host-supplied renderer (a mount function
 * or a custom element, see `ViewerRegionRenderer`) into a container the viewer
 * keeps in the built-in's place, and keeps it current.
 *
 * One {@link RegionSet} owns every region host of one owner (the chrome, or the
 * viewer for dialogs). The shared `ViewerRegionHost` does the change detection
 * and the stable callbacks; the set only decides when to re-render (on every
 * viewer / options change, through {@link RegionSupport.subscribe}) and makes
 * sure each host is destroyed with its container, with the chrome on rebuild
 * and with the viewer on destroy.
 *
 * @module viewer/ui/region-set
 */
import { getRegionRenderer, ViewerRegionHost } from 'ooxml-ui/pptx';
import type {
	HostBackstagePage,
	HostBackstageRegionId,
	ResolvedCustomization,
	ViewerRegionContext,
	ViewerRegionId,
	ViewerRegionRenderer,
} from 'ooxml-ui/pptx';

/** What the viewer lends the region machinery. */
export interface RegionSupport {
	getResolved(): ResolvedCustomization;
	/** The context for `region`; `close` is the region's own dismiss action. */
	buildContext(
		region: ViewerRegionId | HostBackstageRegionId,
		close: () => void,
	): ViewerRegionContext;
	/** Wake `listener` on any viewer or options change; returns the unsubscribe. */
	subscribe(listener: () => void): () => void;
}

export interface RegionHandle {
	/** The container the renderer fills (already attached to the parent). */
	readonly el: HTMLElement;
	/** Destroy the renderer and remove the container. */
	destroy(): void;
}

export interface RegionMountOptions {
	/** The region's dismiss action (default: a no-op). */
	close?: (() => void) | undefined;
}

export interface RegionSet {
	/** True when the host replaced `region` (and did not also hide it). */
	isReplaced(region: ViewerRegionId): boolean;
	/** Mount the host's renderer for `region` into `parent`, or null when it is not replaced. */
	mount(
		region: ViewerRegionId,
		parent: HTMLElement,
		options?: RegionMountOptions,
	): RegionHandle | null;
	/** Mount a host File tab page (`backstage.host.<id>`). */
	mountHostPage(
		page: HostBackstagePage,
		parent: HTMLElement,
		options?: RegionMountOptions,
	): RegionHandle;
	/** Re-render every live host with a fresh context. */
	refresh(): void;
	/**
	 * Detach every live host from the DOM but keep it running, for a chrome
	 * rebuild: a later `mount` of the same region with the same renderer adopts
	 * the parked host (its state survives) instead of building a new one.
	 */
	park(): void;
	/** Destroy the parked hosts nobody adopted (call once the new chrome is built). */
	sweep(): void;
	/** Destroy every live and parked host. */
	destroy(): void;
}

interface Entry {
	region: ViewerRegionId | HostBackstageRegionId;
	renderer: ViewerRegionRenderer;
	host: ViewerRegionHost;
	el: HTMLElement;
	close: () => void;
}

/** A set with no support: nothing is ever replaced (tests, bare chrome). */
export const NO_REGIONS: RegionSet = {
	isReplaced: () => false,
	mount: () => null,
	mountHostPage: (_page, parent) => ({ el: parent, destroy: () => undefined }),
	refresh: () => undefined,
	park: () => undefined,
	sweep: () => undefined,
	destroy: () => undefined,
};

export function createRegionSet(support: RegionSupport | undefined): RegionSet {
	if (!support) {
		return NO_REGIONS;
	}
	const entries = new Set<Entry>();
	// Hosts of a chrome being rebuilt: alive, detached, waiting to be adopted.
	const parked = new Set<Entry>();
	let unsubscribe: (() => void) | null = null;

	const render = (entry: Entry): void => {
		entry.host.render(entry.renderer, support.buildContext(entry.region, entry.close));
	};
	const refresh = (): void => {
		for (const entry of [...entries]) {
			if (!entry.region.startsWith('backstage.host.')) {
				entry.renderer =
					getRegionRenderer(support.getResolved(), entry.region as ViewerRegionId) ??
					entry.renderer;
			}
			render(entry);
		}
	};
	const stopIfIdle = (): void => {
		if (entries.size === 0 && parked.size === 0) {
			unsubscribe?.();
			unsubscribe = null;
		}
	};
	const release = (entry: Entry): void => {
		if (!entries.delete(entry) && !parked.delete(entry)) {
			return;
		}
		entry.host.destroy();
		entry.el.remove();
		stopIfIdle();
	};
	const add = (
		region: Entry['region'],
		renderer: ViewerRegionRenderer,
		parent: HTMLElement,
		options: RegionMountOptions | undefined,
	): RegionHandle => {
		for (const entry of parked) {
			if (entry.region === region && entry.renderer === renderer) {
				// A rebuilt chrome takes the live host back: same container, same state.
				parked.delete(entry);
				entries.add(entry);
				entry.close = options?.close ?? (() => undefined);
				parent.append(entry.el);
				render(entry);
				return { el: entry.el, destroy: () => release(entry) };
			}
		}
		const el = parent.ownerDocument.createElement('div');
		el.className = 'pptxv-region';
		parent.append(el);
		const entry: Entry = {
			region,
			renderer,
			host: new ViewerRegionHost(el),
			el,
			close: options?.close ?? (() => undefined),
		};
		entries.add(entry);
		unsubscribe ??= support.subscribe(refresh);
		render(entry);
		return { el, destroy: () => release(entry) };
	};

	return {
		isReplaced: (region) => getRegionRenderer(support.getResolved(), region) !== undefined,
		mount(region, parent, options) {
			const renderer = getRegionRenderer(support.getResolved(), region);
			return renderer ? add(region, renderer, parent, options) : null;
		},
		mountHostPage: (page, parent, options) =>
			add(`backstage.host.${page.id}`, page.render, parent, options),
		refresh,
		park() {
			for (const entry of entries) {
				entry.el.remove();
				parked.add(entry);
			}
			entries.clear();
		},
		sweep() {
			for (const entry of [...parked]) {
				release(entry);
			}
			stopIfIdle();
		},
		destroy() {
			for (const entry of [...entries, ...parked]) {
				release(entry);
			}
		},
	};
}

/**
 * Hide every built-in child of `frame` and mount the region beside them: the
 * frame (width, border, collapse state) stays, only its content is replaced.
 */
export function replaceFrameContent(
	frame: HTMLElement,
	regions: RegionSet,
	region: ViewerRegionId,
	options?: RegionMountOptions,
): RegionHandle | null {
	const handle = regions.mount(region, frame, options);
	if (handle) {
		frame.classList.add('pptxv-region-replaced');
	}
	return handle;
}
