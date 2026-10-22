/**
 * Host File tab pages (`backstage.extraPages`) in the shared `<office-ui-backstage>`: the nav
 * entries go after the built-in pages of their group, and each page body is a host region
 * (`backstage.host.<id>`) mounted through `HostRegionHost` the first time it is shown. A shown
 * page stays mounted (the File view hides it) until the host removes it, replaces its renderer
 * or the editor is destroyed. Word, Excel and Visio share this; PowerPoint renders its own nav.
 */
import {
	acceptedHostBackstagePages,
	hostBackstageItemId,
	hostBackstagePageId,
	hostBackstageRegionId,
	placeHostBackstagePages,
	type HostBackstagePage,
	type HostBackstageRegionId,
	type HostRegionRenderer,
} from 'ooxml-core/embedding';
import type { OfficeBackstageItem } from '../chrome/backstage';
import { HostRegionHost } from './region-host';

export interface HostBackstagePagesOptions<C> {
	/** The context of a page's region, read every time the page is shown or refreshed. */
	context(
		region: HostBackstageRegionId,
		page: HostBackstagePage<HostRegionRenderer<HTMLElement, C>>,
	): C;
	/** Class of each page container, for products that style their File view pages. */
	pageClass?: string;
	/** Prefix of the error logged when a renderer throws. */
	label?: string;
}

export interface HostBackstagePages<C> {
	/** The nav items with the host pages placed after the built-in pages of their group. */
	items(builtIn: readonly OfficeBackstageItem[]): OfficeBackstageItem[];
	/** Replace the host pages: containers of removed pages are torn down and removed. */
	setPages(
		pages: readonly HostBackstagePage<HostRegionRenderer<HTMLElement, C>>[] | undefined,
	): void;
	/** True when `itemId` is a host page; mounts it (or delivers a new context) when it is. */
	show(itemId: string): boolean;
	/** Deliver a fresh context to every mounted page (after a locale or edit-state change). */
	refresh(): void;
	/** Tear every page down and remove its container. */
	destroy(): void;
}

type Page<C> = HostBackstagePage<HostRegionRenderer<HTMLElement, C>>;

/** Manages the host pages of one File view (`backstage`, an `<office-ui-backstage>`). */
export function createHostBackstagePages<C extends { readonly region: string }>(
	backstage: HTMLElement,
	options: HostBackstagePagesOptions<C>,
): HostBackstagePages<C> {
	let pages: Page<C>[] = [];
	const hosts = new Map<string, HostRegionHost<C>>();

	const container = (page: Page<C>): HostRegionHost<C> => {
		let host = hosts.get(page.id);
		if (host) return host;
		const element = backstage.ownerDocument.createElement('div');
		element.dataset.backstagePage = hostBackstageItemId(page.id);
		if (options.pageClass) element.className = options.pageClass;
		element.hidden = true;
		backstage.append(element);
		host = new HostRegionHost<C>(element, options.label ? { label: options.label } : {});
		hosts.set(page.id, host);
		return host;
	};
	const render = (page: Page<C>, host: HostRegionHost<C>) => {
		const region = hostBackstageRegionId(page.id);
		host.render(page.render, options.context(region, page));
	};
	const drop = (id: string) => {
		const host = hosts.get(id);
		if (!host) return;
		host.destroy();
		host.container.remove();
		hosts.delete(id);
	};

	return {
		items(builtIn) {
			return placeHostBackstagePages(builtIn, pages, (page) => ({
				id: hostBackstageItemId(page.id),
				label: page.label,
				...(page.group === 'footer' ? { group: 'footer' as const } : {}),
			}));
		},
		setPages(next) {
			pages = acceptedHostBackstagePages(next);
			const kept = new Set(pages.map((page) => page.id));
			for (const id of [...hosts.keys()]) if (!kept.has(id)) drop(id);
			// A mounted page whose renderer changed remounts now, so it never shows stale content.
			for (const page of pages) {
				const host = hosts.get(page.id);
				if (host?.mounted) render(page, host);
			}
		},
		show(itemId) {
			const id = hostBackstagePageId(itemId);
			const page = id === undefined ? undefined : pages.find((candidate) => candidate.id === id);
			if (!page) return false;
			render(page, container(page));
			return true;
		},
		refresh() {
			for (const page of pages) {
				const host = hosts.get(page.id);
				if (host?.mounted) render(page, host);
			}
		},
		destroy() {
			for (const id of [...hosts.keys()]) drop(id);
		},
	};
}
