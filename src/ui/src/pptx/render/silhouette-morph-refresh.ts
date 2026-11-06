/**
 * Brings silhouette filters that are already painted up to date with a morph.
 *
 * `getSilhouetteEffects` adds the outgoing end's layers to the filter of the
 * shape that arrives, but only for markup built after the plan registered the
 * morph. A binding that painted the stage first (memoized or cloned markup)
 * calls {@link refreshSilhouetteMorphFilters} once the plan is built.
 *
 * @module render/silhouette-morph-refresh
 */
import { buildSilhouetteFilterMarkup, MORPH_PROGRESS_VAR } from './silhouette-effects';
import { silhouetteMorphEntry } from './silhouette-morph-registry';

const SVG_NS = 'http://www.w3.org/2000/svg';

/**
 * Rebuilds, under `root`, each silhouette filter whose shape a registered morph
 * brings in. A filter that already carries the dissolve layers is left alone.
 */
export function refreshSilhouetteMorphFilters(root: ParentNode): void {
	for (const filter of root.querySelectorAll('filter[id^="shape-fx-"]')) {
		const id =
			filter.closest('[data-pptx-morph-fx]')?.getAttribute('data-pptx-morph-fx') ??
			filter.closest('[data-element-id]')?.getAttribute('data-element-id');
		const entry = id ? silhouetteMorphEntry(id) : undefined;
		if (!entry || filter.innerHTML.includes(MORPH_PROGRESS_VAR)) {
			continue;
		}
		const holder = filter.ownerDocument.createElementNS(SVG_NS, 'svg');
		holder.innerHTML = buildSilhouetteFilterMarkup(
			filter.id,
			entry.width,
			entry.height,
			entry.params,
			entry.out,
		);
		const fresh = holder.firstElementChild;
		if (!fresh) {
			continue;
		}
		const original = filter.getAttribute('data-pptx-sid');
		if (original) {
			fresh.setAttribute('data-pptx-sid', original);
		}
		filter.replaceWith(fresh);
	}
}
