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
import {
	rewriteCustomShowAction as rewriteAction,
	type CustomShowActionRewrite,
} from './custom-show-link';
import { mapCustomShowTextLinks } from './custom-show-text-links';
import { mapCustomShowTableLinks } from './custom-show-table-links';

export { customShowActionId } from './custom-show-link';
export type { CustomShowActionRewrite } from './custom-show-link';

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

/** Updates each element once while keeping default/run link mirrors consistent. */
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
	let clicks = 0;
	let hovers = 0;
	const segments = 'textSegments' in element ? element.textSegments : undefined;
	if (segments) {
		const mapped = segments.map((segment) => {
			if (!segment.style) return segment;
			const result = mapCustomShowTextLinks(segment.style, rewrite);
			clicks += Number(result.click);
			hovers += Number(result.hover);
			return result.style === segment.style ? segment : { ...segment, style: result.style };
		});
		if (clicks || hovers) edit().textSegments = mapped;
	}
	changed += clicks + hovers;
	const textStyle = 'textStyle' in element ? element.textStyle : undefined;
	if (textStyle) {
		const result = mapCustomShowTextLinks(textStyle, rewrite);
		if (result.style !== textStyle) edit().textStyle = result.style;
		// The default text style mirrors run metadata; do not count that copy twice.
		if (result.click && clicks === 0) changed++;
		if (result.hover && hovers === 0) changed++;
	}
	if (element.type === 'table' && element.rawXml) {
		const result = mapCustomShowTableLinks(element.rawXml, rewrite);
		if (result.changed) edit().rawXml = result.rawXml;
		changed += result.changed;
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
