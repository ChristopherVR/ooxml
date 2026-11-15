/**
 * The DOM-free half of drawing one element of a styled stage clone on its own (see
 * `pdf-vector-raster-batch.ts`). `rasterizeIsolated` hides the live stage, shows the target (and
 * hides the target's children when it draws only the target's own box or text shadow), and lets
 * the cascade work out every other node. A clone with every computed style inlined has no
 * inheritance left to work that out, so the same answer is computed here for every element and
 * written into the serialised clone explicitly. A subtree the picture shows none of is left out
 * of its markup where that cannot move anything (see `pdf-vector-raster-template.ts`).
 */
import type { IsolatedRasterMode } from './pdf-vector-raster';

/** Elements in document order (a parent before its children), as a styled clone holds them. */
export interface IsolatedVisibilityTree {
	/** Index of each element's parent element, `-1` for the root (the stage). */
	parents: readonly number[];
	/** Each element's own computed `visibility` as inlined on the clone (`''` when it has none). */
	visibility: readonly string[];
	/**
	 * The `visibility` each live element declares in its own `style` attribute (`''` for none):
	 * a declared value beats inheritance, so it holds while the stage around it is hidden.
	 */
	declared: readonly string[];
}

/**
 * The serialised clone as a template: `chunks[k]` precedes `ops[k]`, and the last chunk follows
 * the last op. A `style` op is a slot at the end of element `style`'s style attribute; `open` and
 * `close` bracket the markup of an element whose subtree may be left out.
 */
export interface IsolatedTemplate extends IsolatedVisibilityTree {
	/** Whether element `i`'s subtree may be left out of a picture that shows none of it. */
	prunable: readonly boolean[];
	chunks: readonly string[];
	ops: readonly IsolatedTemplateOp[];
}

export type IsolatedTemplateOp =
	| { kind: 'style'; element: number }
	/** `end` is the index of the matching `close` op. */
	| { kind: 'open'; element: number; end: number }
	| { kind: 'close'; element: number };

/** What draws only the shadow of an element's own text (as `rasterizeIsolated` sets it). */
export const SHADOW_ONLY_DECLARATIONS: readonly string[] = [
	'color:transparent',
	'-webkit-text-fill-color:transparent',
	'background-color:transparent',
	'background-image:none',
	'border-top-color:transparent',
	'border-right-color:transparent',
	'border-bottom-color:transparent',
	'border-left-color:transparent',
	'box-shadow:none',
	'text-decoration-color:transparent',
];

const VISIBILITY = new Set(['visible', 'hidden', 'collapse']);

/**
 * The `visibility` of every element when only element `target` is drawn, `mode` as
 * `rasterizeIsolated` draws it.
 *
 * After the stage, the target and (for `self` and `shadow`) its children, which the capture sets
 * itself, an element keeps a value its own `style` declares; then an element whose inlined value
 * differs from its parent's was given its own value by a stylesheet and keeps it too; any other
 * element inherits its parent's new value.
 */
export function isolatedVisibilities(
	tree: IsolatedVisibilityTree,
	target: number,
	mode: IsolatedRasterMode,
): string[] {
	const { parents, visibility, declared } = tree;
	const out: string[] = [];
	for (let i = 0; i < parents.length; i++) {
		const parent = parents[i] ?? -1;
		const own = visibility[i] ?? '';
		const set = declared[i] ?? '';
		let value: string;
		if (i === target) value = 'visible';
		else if (parent === target && mode !== 'subtree') value = 'hidden';
		else if (parent < 0) value = 'hidden';
		else if (VISIBILITY.has(set)) value = set;
		else if (own && own !== (visibility[parent] ?? '')) value = own;
		else value = out[parent] ?? 'visible';
		out.push(value);
	}
	return out;
}

/** Whether each element or anything inside it is shown. */
function shownSubtrees(parents: readonly number[], shown: readonly string[]): boolean[] {
	const any = shown.map((value) => value === 'visible');
	for (let i = any.length - 1; i > 0; i--) {
		const parent = parents[i] ?? -1;
		if (any[i] && parent >= 0) any[parent] = true;
	}
	return any;
}

export interface IsolatedMarkupOptions {
	/** `false` keeps every subtree (to check that leaving them out changes nothing). */
	prune?: boolean;
	/**
	 * The declarations a `shadow` picture gives its target, `!important` each; the capture's own
	 * overrides ({@link SHADOW_ONLY_DECLARATIONS}) when omitted.
	 */
	shadowStyle?: string;
}

/** The serialised clone with only `target` shown the way `mode` draws it. */
export function isolatedMarkup(
	template: IsolatedTemplate,
	target: number,
	mode: IsolatedRasterMode,
	options: IsolatedMarkupOptions = {},
): string {
	const { prune = true } = options;
	const shown = isolatedVisibilities(template, target, mode);
	const anyShown = shownSubtrees(template.parents, shown);
	const leaveOut = (element: number) =>
		prune && template.prunable[element] === true && !anyShown[element];
	const shadow =
		options.shadowStyle ??
		SHADOW_ONLY_DECLARATIONS.map((declaration) => `;${declaration}!important`).join('');
	const { chunks, ops } = template;
	const out: string[] = [];
	for (let k = 0; k < ops.length; k++) {
		out.push(chunks[k] ?? '');
		const op = ops[k]!;
		if (op.kind === 'style') {
			out.push(`visibility:${shown[op.element]}!important`);
			if (op.element === target && mode === 'shadow') out.push(shadow);
		} else if (op.kind === 'open' && leaveOut(op.element)) {
			// Skip to the matching close: its chunk (the subtree's own markup) is left out too.
			k = op.end;
		}
	}
	out.push(chunks[ops.length] ?? '');
	return out.join('');
}
