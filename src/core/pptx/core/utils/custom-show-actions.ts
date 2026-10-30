/**
 * Custom show ids and the actions that name them.
 *
 * `p:custShow/@id` is an `xsd:unsignedInt`, and `ppaction://customshow?id=N` (shape actions,
 * text hyperlinks) and `p:showPr/p:custShow/@id` refer to it. These helpers keep the three in
 * step: {@link normalizeCustomShowIds} makes every id a valid unique unsigned integer, and
 * {@link mapCustomShowActions} rewrites or clears the actions that point at an id.
 *
 * @module utils/custom-show-actions
 */

import type { PptxCustomShow, PptxElement, PptxSlide } from '../types';

const UNSIGNED_INT_MAX = 4294967295;

function isUnsignedInt(id: string): boolean {
	return /^\d+$/.test(id) && Number(id) <= UNSIGNED_INT_MAX;
}

/**
 * Give every show a valid, unique `xsd:unsignedInt` id. Valid ids keep their value; an invalid
 * or duplicate one takes the next free number. `idMap` holds only the ids that changed, keyed by
 * the old id (the first show with a duplicated old id keeps it, so the map never names it).
 */
export function normalizeCustomShowIds(shows: readonly PptxCustomShow[]): {
	shows: PptxCustomShow[];
	idMap: Map<string, string>;
} {
	const used = new Set<string>();
	let max = -1;
	for (const show of shows) {
		if (isUnsignedInt(show.id) && !used.has(show.id)) {
			used.add(show.id);
			max = Math.max(max, Number(show.id));
		}
	}
	const kept = new Set<string>();
	const idMap = new Map<string, string>();
	const next = shows.map((show) => {
		if (isUnsignedInt(show.id) && !kept.has(show.id)) {
			kept.add(show.id);
			return show;
		}
		const id = String(++max);
		if (!idMap.has(show.id)) {
			idMap.set(show.id, id);
		}
		// `rawXml` still carries the old id attribute; the writer sets the id from `show.id`.
		return { ...show, id };
	});
	return {
		shows: idMap.size > 0 || next.some((s, i) => s !== shows[i]) ? next : [...shows],
		idMap,
	};
}

/** Reads the `id` query value of a `ppaction://customshow?...` action, else undefined. */
export function customShowActionId(action: string | undefined): string | undefined {
	if (!action || !action.toLowerCase().includes('ppaction://customshow')) {
		return undefined;
	}
	const match = action.match(/[?&]id=([^&]*)/i);
	return match ? decodeURIComponent(match[1] as string) : undefined;
}

/** What to do with an action that jumps into custom show `id`: a new id, or null to remove it. */
export type CustomShowActionRewrite = (id: string) => string | null;

function rewriteAction(action: string | undefined, rewrite: CustomShowActionRewrite) {
	const id = customShowActionId(action);
	if (id === undefined) {
		return { changed: false, action };
	}
	const next = rewrite(id);
	if (next === id) {
		return { changed: false, action };
	}
	return {
		changed: true,
		action:
			next === null
				? undefined
				: (action as string).replace(/([?&]id=)[^&]*/i, `$1${encodeURIComponent(next)}`),
	};
}

function mapElement(
	element: PptxElement,
	rewrite: CustomShowActionRewrite,
): { element: PptxElement; changed: number } {
	let changed = 0;
	let next: PptxElement = element;
	const edit = (): Record<string, unknown> => {
		if (next === element) {
			next = { ...element };
		}
		return next as unknown as Record<string, unknown>;
	};
	for (const key of ['actionClick', 'actionHover'] as const) {
		const current = element[key];
		const result = rewriteAction(current?.action, rewrite);
		if (current && result.changed) {
			changed++;
			if (result.action === undefined) {
				delete edit()[key];
			} else {
				edit()[key] = { ...current, action: result.action };
			}
		}
	}
	const segments = 'textSegments' in element ? element.textSegments : undefined;
	if (segments?.some((s) => customShowActionId(s.style?.hyperlinkAction) !== undefined)) {
		edit().textSegments = segments.map((segment) => {
			const result = rewriteAction(segment.style?.hyperlinkAction, rewrite);
			if (!result.changed) {
				return segment;
			}
			changed++;
			if (result.action !== undefined) {
				return { ...segment, style: { ...segment.style, hyperlinkAction: result.action } };
			}
			const {
				hyperlink: _h,
				hyperlinkRId: _r,
				hyperlinkTooltip: _t,
				hyperlinkAction: _a,
				hyperlinkTargetSlideIndex: _i,
				...style
			} = segment.style;
			return { ...segment, style };
		});
	}
	if (element.type === 'group') {
		const children = element.children.map((child) => mapElement(child, rewrite));
		if (children.some((c) => c.changed > 0)) {
			edit().children = children.map((c) => c.element);
			changed += children.reduce((n, c) => n + c.changed, 0);
		}
	}
	return { element: next, changed };
}

/**
 * Rewrite or remove every custom-show jump (shape click/hover actions and text hyperlinks,
 * groups included). Returns new slides; slides with nothing to change keep their identity.
 */
export function mapCustomShowActions(
	slides: readonly PptxSlide[],
	rewrite: CustomShowActionRewrite,
): { slides: PptxSlide[]; changed: number } {
	let changed = 0;
	const next = slides.map((slide) => {
		const results = slide.elements.map((element) => mapElement(element, rewrite));
		const n = results.reduce((sum, r) => sum + r.changed, 0);
		if (n === 0) {
			return slide;
		}
		changed += n;
		return { ...slide, elements: results.map((r) => r.element) };
	});
	return { slides: next, changed };
}
