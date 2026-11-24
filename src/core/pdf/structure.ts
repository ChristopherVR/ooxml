/**
 * The structure tree of a tagged PDF (ISO 32000-1 section 14.7): each marked-content id of a page
 * mapped to the block element it belongs to (its standard type after the role map, and an id per
 * element), so text keeps its paragraphs and headings and pictures their alternate text. Read
 * within node and depth limits; files without a tree, or with a damaged one, give no marks.
 */
import type { PdfDocument } from './document';
import {
	decodeTextString,
	isDict,
	isName,
	isNumber,
	PdfRef,
	type PdfDict,
	type PdfValue,
} from './objects';

export interface StructureMark {
	/** The standard type of the block element (`P`, `H1`, `LI`, `TD`, `Figure`...). */
	role: string;
	/** An id shared by the marks of one block element. */
	element: number;
	/** Alternate text of the nearest element that has some (a figure's description). */
	alt?: string;
}

/** Page index to marked-content id to mark. */
export type StructureMarks = Map<number, Map<number, StructureMark>>;

const MAX_NODES = 200_000;
const MAX_DEPTH = 64;

const STANDARD = new Set(
	'Document DocumentFragment Part Art Sect Div BlockQuote Caption TOC TOCI Index NonStruct Private Aside Title FENote P H H1 H2 H3 H4 H5 H6 L LI Lbl LBody Table TR TH TD THead TBody TFoot Span Quote Note Reference BibEntry Code Link Annot Ruby RB RT RP Warichu WT WP Figure Formula Form Sub Em Strong'.split(
		' ',
	),
);
/** Types that sit inside a block (a list item's label and body belong to the item). */
const INLINE = new Set(
	'Span Quote Reference BibEntry Code Link Annot Ruby RB RT RP Warichu WT WP Lbl LBody Sub Em Strong NonStruct'.split(
		' ',
	),
);

function standardRole(raw: string, roleMap: PdfDict | undefined): string {
	let role = raw;
	for (let step = 0; step < 8 && !STANDARD.has(role); step++) {
		const mapped = roleMap?.get(role);
		if (!isName(mapped)) break;
		role = mapped.name;
	}
	return role;
}

interface Context {
	page: number | undefined;
	role: string;
	element: number;
	alt: string | undefined;
}

/** Reads the marks of a tagged file's structure tree (empty when it has none). */
export async function readStructure(
	document: PdfDocument,
	catalog: PdfDict | undefined,
	pageRefs: Map<string, number>,
): Promise<StructureMarks> {
	const marks: StructureMarks = new Map();
	const root = await document.dict(catalog?.get('StructTreeRoot'));
	if (!root) return marks;
	const roleMap = await document.dict(root.get('RoleMap'));
	const seen = new Set<string>();
	let nodes = 0;
	let elements = 0;
	const pageOf = (value: PdfValue | undefined, fallback: number | undefined) =>
		value instanceof PdfRef ? (pageRefs.get(value.key) ?? fallback) : fallback;
	const mark = (page: number | undefined, mcid: number, context: Context) => {
		if (page === undefined) return;
		let map = marks.get(page);
		if (!map) marks.set(page, (map = new Map()));
		map.set(mcid, {
			role: context.role,
			element: context.element,
			...(context.alt ? { alt: context.alt } : {}),
		});
	};
	const visit = async (value: PdfValue | undefined, context: Context, depth: number) => {
		if (depth > MAX_DEPTH || ++nodes > MAX_NODES) return;
		if (isNumber(value)) return mark(context.page, value, context);
		if (value instanceof PdfRef) {
			if (seen.has(value.key)) return;
			seen.add(value.key);
		}
		const resolved = await document.resolve(value);
		if (Array.isArray(resolved)) {
			for (const kid of resolved) await visit(kid, context, depth + 1);
			return;
		}
		if (!isDict(resolved)) return;
		const type = resolved.get('Type');
		if (isName(type, 'MCR')) {
			const mcid = await document.resolve(resolved.get('MCID'));
			if (isNumber(mcid)) mark(pageOf(resolved.get('Pg'), context.page), mcid, context);
			return;
		}
		if (isName(type, 'OBJR')) return;
		const raw = resolved.get('S');
		const page = pageOf(resolved.get('Pg'), context.page);
		let next: Context = { ...context, page };
		if (isName(raw)) {
			const role = standardRole(raw.name, roleMap);
			if (!INLINE.has(role) || context.element < 0) next = { ...next, role, element: elements++ };
		}
		const alt = decodeTextString(await document.resolve(resolved.get('Alt')));
		if (alt?.trim()) next.alt = alt.trim();
		await visit(resolved.get('K'), next, depth + 1);
	};
	await visit(root.get('K'), { page: undefined, role: 'Document', element: -1, alt: undefined }, 0);
	return marks;
}
