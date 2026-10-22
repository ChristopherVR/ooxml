/**
 * Host File tab (Backstage) pages: pages a host adds to an editor's File tab, after the
 * built-in pages of their group. The page body is a host region (`backstage.host.<id>`), so it
 * takes any renderer form of `HostRegionRenderer`; the navigation entry is plain data.
 */

/** Where a page's navigation entry sits: with the main pages, or next to Account and Options. */
export type HostBackstageGroup = 'main' | 'footer';

/**
 * One host page of the File tab. `Renderer` is the product's region renderer type (a mount
 * function over an `HTMLElement`, or a custom element), bound by each product.
 */
export interface HostBackstagePage<Renderer = unknown> {
	/** Unique among host pages; the region id is `backstage.host.<id>`. A repeat is ignored. */
	id: string;
	/** The navigation label. The host localises it; the editor shows it as given. */
	label: string;
	/** `footer` places it next to Account and Options; default `main`. */
	group?: HostBackstageGroup;
	/**
	 * Sort key among the host pages of the same group (ascending, default 0). Host pages always
	 * come after the built-in pages of their group; equal keys keep the order they were given.
	 */
	order?: number;
	/** The page body. */
	render: Renderer;
}

/** The nav id of host page `id`, kept apart from every built-in page id. */
export function hostBackstageItemId(id: string): `host.${string}` {
	return `host.${id}`;
}

/** The host page id of a nav id made by `hostBackstageItemId`, or undefined for a built-in. */
export function hostBackstagePageId(itemId: string): string | undefined {
	return itemId.startsWith('host.') ? itemId.slice(5) : undefined;
}

/**
 * The host pages an editor shows, in display order: pages with an empty or repeated id are left
 * out (the first one wins), then the rest sort by `order` (stable, so equal keys keep the host's
 * order).
 */
export function acceptedHostBackstagePages<P extends HostBackstagePage<unknown>>(
	pages: readonly P[] | null | undefined,
): P[] {
	const seen = new Set<string>();
	const accepted: Array<{ page: P; index: number }> = [];
	for (const page of pages ?? []) {
		if (!page || typeof page.id !== 'string' || !page.id || seen.has(page.id)) continue;
		seen.add(page.id);
		accepted.push({ page, index: accepted.length });
	}
	return accepted
		.sort((a, b) => (a.page.order ?? 0) - (b.page.order ?? 0) || a.index - b.index)
		.map(({ page }) => page);
}

/** A nav entry with an optional group: what `placeHostBackstagePages` interleaves. */
export interface BackstageNavSlot {
	id: string;
	group?: HostBackstageGroup | undefined;
}

/**
 * The File tab navigation with host pages placed: each host entry goes after the last built-in
 * entry of its group (at the end when its group has none). `toEntry` makes the product's nav
 * entry for a page; `items` is not changed.
 */
export function placeHostBackstagePages<
	T extends BackstageNavSlot,
	P extends HostBackstagePage<unknown>,
>(items: readonly T[], pages: readonly P[] | null | undefined, toEntry: (page: P) => T): T[] {
	const entries = [...items];
	for (const page of acceptedHostBackstagePages(pages)) {
		const group = page.group ?? 'main';
		let at = entries.length;
		while (at > 0 && (entries[at - 1]?.group ?? 'main') !== group) at--;
		entries.splice(at === 0 ? entries.length : at, 0, toEntry(page));
	}
	return entries;
}
