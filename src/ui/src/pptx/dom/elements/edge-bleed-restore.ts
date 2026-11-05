/**
 * Undoing the abutting-edge bleed (`abutting-edge-bleed.ts`) on rendered DOM.
 *
 * A grown box is only hidden by the cover that stands on top of it in the
 * stage. A morph transition overlay clones the outgoing stage and moves its
 * shapes apart, so a clone of the grown box would carry its 2 px growth into
 * the open. The renderer therefore records the authored value of every style
 * key the bleed patch overwrote, and {@link stripAbuttingEdgeBleeds} puts them
 * back on a clone before it is animated.
 */
import type { CssStyleMap } from '../../index';

const ATTRIBUTE = 'data-pptx-edge-bleed';

const kebab = (key: string): string => key.replace(/[A-Z]/gu, (c) => `-${c.toLowerCase()}`);

/**
 * Remember, on `el`, the value each key of `patch` had in `before` (the style
 * the element was painted with ahead of the bleed). A no-op for an empty patch.
 */
export function markAbuttingEdgeBleed(
	el: HTMLElement,
	patch: CssStyleMap,
	before: Readonly<Record<string, unknown>>,
): void {
	const keys = Object.keys(patch);
	if (keys.length === 0) return;
	const original: Record<string, string> = {};
	for (const key of keys) {
		const value = before[key] ?? before[kebab(key)];
		original[kebab(key)] = value === undefined || value === null ? '' : String(value);
	}
	el.setAttribute(ATTRIBUTE, JSON.stringify(original));
}

/** Restore the authored box and fill of every grown element under `root`. */
export function stripAbuttingEdgeBleeds(root: ParentNode): void {
	const nodes = [...root.querySelectorAll<HTMLElement>(`[${ATTRIBUTE}]`)];
	if ('hasAttribute' in root && (root as Element).hasAttribute(ATTRIBUTE)) {
		nodes.push(root as HTMLElement);
	}
	for (const node of nodes) {
		try {
			const original = JSON.parse(node.getAttribute(ATTRIBUTE) ?? '{}') as Record<string, string>;
			for (const [property, value] of Object.entries(original)) {
				if (value === '') node.style.removeProperty(property);
				else node.style.setProperty(property, value);
			}
		} catch {
			// A malformed marker leaves the clone as it was.
		}
		node.removeAttribute(ATTRIBUTE);
	}
}
