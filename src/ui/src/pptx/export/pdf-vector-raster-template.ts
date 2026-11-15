/**
 * Serialise a styled stage clone once as an {@link IsolatedTemplate}: every element gets a slot
 * at the end of its `style` attribute (where a picture writes its `visibility`), and each
 * subtree a picture may leave out is bracketed by marker comments. The markers and slots are
 * found again by a per-template nonce no slide text can contain, and none of them reach a
 * picture: each slot is replaced and each marker dropped, so a picture's markup is the clone
 * serialised as `XMLSerializer` writes it.
 *
 * A subtree may be left out of a picture that shows none of it only when leaving it out cannot
 * change anything else's pixels: it is out of flow (`position: absolute` or `fixed`, so it never
 * moves a sibling), its parent does not scroll (so no scrollbar depends on it), nothing in it
 * sets a CSS counter, and nothing outside it refers to an `id` inside it (a gradient, filter,
 * clip path or `<use>` target).
 */
import type { IsolatedTemplate, IsolatedTemplateOp } from './pdf-vector-raster-isolate';
import { RASTER_VISIBILITY_ATTRIBUTE } from './pdf-vector-raster-live';

const OUT_OF_FLOW = new Set(['absolute', 'fixed']);
const SCROLLS = new Set(['auto', 'scroll', 'overlay']);
const COUNTERS = ['counter-reset', 'counter-increment', 'counter-set'];

function styleValue(el: Element, property: string): string {
	return (el as HTMLElement).style?.getPropertyValue(property) ?? '';
}

/** Whether `markup` refers (`#id`) to an id inside `subtree` anywhere outside `[start, end)`. */
function referencedFromOutside(
	subtree: Element,
	markup: string,
	start: number,
	end: number,
): boolean {
	const ids = [subtree, ...subtree.querySelectorAll('[id]')]
		.map((el) => el.getAttribute('id'))
		.filter((id): id is string => !!id);
	for (const id of ids) {
		const needle = `#${id}`;
		for (let at = markup.indexOf(needle); at >= 0; at = markup.indexOf(needle, at + 1))
			if (at < start || at >= end) return true;
	}
	return false;
}

/**
 * Build the template of `clone` (detached, styles inlined). `targetAttribute` marks the elements
 * to draw: the result maps each mark's value to the element's index. That mark and
 * {@link RASTER_VISIBILITY_ATTRIBUTE} are read and removed. The clone is changed (slots,
 * markers) and must not be reused.
 */
export function buildIsolatedTemplate(
	clone: Element,
	targetAttribute: string,
): { template: IsolatedTemplate; targets: Map<string, number> } {
	const doc = clone.ownerDocument;
	const elements: Element[] = [clone, ...clone.querySelectorAll('*')];
	const index = new Map<Element, number>(elements.map((el, i) => [el, i]));
	const parents: number[] = [];
	const visibility: string[] = [];
	const declared: string[] = [];
	const counted: boolean[] = [];
	const targets = new Map<string, number>();
	elements.forEach((el, i) => {
		parents.push(el.parentElement ? (index.get(el.parentElement) ?? -1) : -1);
		visibility.push(styleValue(el, 'visibility'));
		declared.push(el.getAttribute(RASTER_VISIBILITY_ATTRIBUTE) ?? '');
		el.removeAttribute(RASTER_VISIBILITY_ATTRIBUTE);
		counted.push(COUNTERS.some((property) => !['', 'none'].includes(styleValue(el, property))));
		const tag = el.getAttribute(targetAttribute);
		if (tag !== null) {
			targets.set(tag, i);
			el.removeAttribute(targetAttribute);
		}
	});
	for (let i = elements.length - 1; i > 0; i--) if (counted[i]) counted[parents[i]!] = true;

	const nonce = `p${Math.random().toString(36).slice(2)}`;
	const candidates = new Set<number>();
	elements.forEach((el, i) => {
		const parent = el.parentElement;
		if (!parent || i === 0 || counted[i] || !OUT_OF_FLOW.has(styleValue(el, 'position'))) return;
		if (
			SCROLLS.has(styleValue(parent, 'overflow-x')) ||
			SCROLLS.has(styleValue(parent, 'overflow-y'))
		)
			return;
		candidates.add(i);
		parent.insertBefore(doc.createComment(`${nonce}b${i}`), el);
		parent.insertBefore(doc.createComment(`${nonce}e${i}`), el.nextSibling);
	});
	elements.forEach((el, i) => {
		const own = el.getAttribute('style');
		el.setAttribute('style', `${own ? `${own};` : ''}--${nonce}-${i}:0`);
	});

	const markup = new XMLSerializer().serializeToString(clone);
	const chunks: string[] = [];
	const ops: IsolatedTemplateOp[] = [];
	const opened: { op: number; start: number }[] = [];
	const ranges = new Map<number, [number, number]>();
	const marker = new RegExp(`--${nonce}-(\\d+):0|<!--${nonce}([be])(\\d+)-->`, 'gu');
	let at = 0;
	for (const match of markup.matchAll(marker)) {
		const found = match.index ?? 0;
		chunks.push(markup.slice(at, found));
		at = found + match[0].length;
		if (match[1] !== undefined) {
			ops.push({ kind: 'style', element: Number(match[1]) });
			continue;
		}
		const element = Number(match[3]);
		if (match[2] === 'b') {
			opened.push({ op: ops.length, start: at });
			ops.push({ kind: 'open', element, end: -1 });
			continue;
		}
		const open = opened.pop();
		const openOp = open && ops[open.op];
		if (!open || openOp?.kind !== 'open' || openOp.element !== element)
			throw new Error('The clone markers are out of order.');
		openOp.end = ops.length;
		ranges.set(element, [open.start, found]);
		ops.push({ kind: 'close', element });
	}
	chunks.push(markup.slice(at));
	if (ops.filter((op) => op.kind === 'style').length !== elements.length)
		throw new Error('A clone element lost its style slot.');

	const prunable = elements.map((el, i) => {
		const range = ranges.get(i);
		return !!range && candidates.has(i) && !referencedFromOutside(el, markup, range[0], range[1]);
	});
	return {
		template: { parents, visibility, declared, prunable, chunks, ops },
		targets,
	};
}
