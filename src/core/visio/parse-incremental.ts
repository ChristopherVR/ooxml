// Re-reads only some pages of a package whose other parts did not change. `parseVsdx(bytes,
// { pages })` reads the document-level data (styles, themes, masters) and just the named pages;
// `mergeVisioPages` splices those pages into the document read earlier, so unchanged pages are
// reused as the same objects. The merge returns undefined whenever the result could differ from a
// whole read (a limit that applies across pages, diagnostics that cannot be placed), and the
// caller then reads the whole package.
import type { VisioDocument, VisioDiagnostic, VisioShape } from './model';
import { parseVsdx, type ParseVsdxOptions } from './parser';

/** Diagnostics that validating the page list adds; they depend on no page part. */
const BACKGROUND_CODES = new Set(['background-cycle', 'missing-background-page']);
/** Above this many diagnostics the collector's limits could make a spliced list differ. */
const MAX_SPLICED_DIAGNOSTICS = 1000;
const DEFAULT_MAX_SHAPES = 25_000;

const countShapes = (shapes: readonly VisioShape[]): number =>
	shapes.reduce((sum, shape) => sum + 1 + countShapes(shape.children), 0);

/** The diagnostics a page's own read produced (not the list-level background checks). */
const isPageDiagnostic = (d: VisioDiagnostic): boolean =>
	d.pageId !== undefined && !BACKGROUND_CODES.has(d.code);

/**
 * `previous` with the pages of `partial` (read with `pages`) in place of the same pages. Only
 * valid when the pages part, the masters, styles, theme, media and every other part are the ones
 * `previous` was read from. Returns undefined when the merge cannot match a whole read.
 */
export function mergeVisioPages(
	previous: VisioDocument,
	partial: VisioDocument,
	pageIds: readonly string[],
	options: { maxShapes?: number } = {},
): VisioDocument | undefined {
	if (previous.format !== 'vsdx' || partial.format !== 'vsdx') return undefined;
	const wanted = new Set(pageIds);
	const fresh = new Map(partial.pages.map((page) => [page.id, page]));
	if (fresh.size !== wanted.size || ![...wanted].every((id) => fresh.has(id))) return undefined;
	const before = new Map(previous.pages.map((page) => [page.id, page]));
	if (![...wanted].every((id) => before.has(id))) return undefined;
	// Background validation removes a reference it cannot resolve; the page list did not change,
	// so the earlier outcome still holds.
	for (const [id, page] of fresh)
		if (before.get(id)!.backgroundPageId === undefined) delete page.backgroundPageId;
	const pages = previous.pages.map((page) => fresh.get(page.id) ?? page);
	if (
		pages.reduce((sum, page) => sum + countShapes(page.shapes), 0) >
		(options.maxShapes ?? DEFAULT_MAX_SHAPES)
	)
		return undefined;
	const diagnostics = spliceDiagnostics(previous.diagnostics, partial.diagnostics, wanted);
	if (!diagnostics) return undefined;
	return { ...previous, pages, diagnostics };
}

function spliceDiagnostics(
	previous: readonly VisioDiagnostic[],
	partial: readonly VisioDiagnostic[],
	wanted: ReadonlySet<string>,
): VisioDiagnostic[] | undefined {
	const truncated = (list: readonly VisioDiagnostic[]) =>
		list.some((d) => d.code === 'diagnostics-truncated');
	if (truncated(previous) || truncated(partial)) return undefined;
	const fresh = partial.filter((d) => isPageDiagnostic(d) && wanted.has(d.pageId!));
	if (previous.length + fresh.length > MAX_SPLICED_DIAGNOSTICS) return undefined;
	let out = [...previous];
	for (const id of wanted) {
		const next = fresh.filter((d) => d.pageId === id);
		const at = out.flatMap((d, i) => (isPageDiagnostic(d) && d.pageId === id ? [i] : []));
		if (!at.length) {
			if (next.length) return undefined; // Where they belong among the list-level ones is unknown.
			continue;
		}
		const first = at[0]!;
		const last = at[at.length - 1]!;
		if (last - first + 1 !== at.length) return undefined;
		out = [...out.slice(0, first), ...next, ...out.slice(last + 1)];
	}
	return out;
}

/**
 * Reads the pages `pageIds` of `input` and merges them into `previous`; undefined means read the
 * whole package. Never throws for a merge that does not apply (a parse error still propagates).
 */
export async function parseVsdxPages(
	input: Uint8Array | ArrayBuffer,
	previous: VisioDocument,
	pageIds: readonly string[],
	options: ParseVsdxOptions = {},
): Promise<VisioDocument | undefined> {
	const partial = await parseVsdx(input, { ...options, pages: pageIds });
	return mergeVisioPages(previous, partial, pageIds, {
		...(options.maxShapes !== undefined ? { maxShapes: options.maxShapes } : {}),
	});
}
