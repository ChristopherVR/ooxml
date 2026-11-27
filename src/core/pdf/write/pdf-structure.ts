/**
 * Navigation and accessibility objects for `buildVectorPdf`: link and note annotations, the
 * document outline (bookmarks) and the logical structure tree of a tagged PDF (ISO 32000-1
 * 12.3.3, 12.5, 14.7). Callers describe them with page objects and marked-content ids; the
 * writer turns them into indirect objects. Every part is optional, so pages that set none (the
 * PowerPoint and Visio exports) are written as before. DOM-free.
 */
import { pdfLiteral } from './pdf-fonts';
import type { PdfVectorPage } from './pdf-vector';

/** A place on a page, in PDF points with y up from the page's bottom edge. */
export interface PdfDestination {
	page: PdfVectorPage;
	/** Left edge shown; the viewer keeps its scroll position when unset. */
	x?: number;
	/** Top edge shown. */
	y: number;
}

/** A clickable area that opens a URI or goes to a place in the document. */
export interface PdfLinkAnnotation {
	kind: 'link';
	/** `[left, bottom, right, top]` in points, y up. */
	rect: readonly [number, number, number, number];
	uri?: string;
	dest?: PdfDestination;
	/** Alternate description, read by assistive technology. */
	contents?: string;
}

/** A sticky note (a comment), shown as an icon at `rect`. */
export interface PdfNoteAnnotation {
	kind: 'note';
	rect: readonly [number, number, number, number];
	contents: string;
	author?: string;
	/** Modification date as a PDF date string (`D:YYYYMMDDHHmmSSZ`) or ISO date. */
	date?: string;
}

export type PdfAnnotation = PdfLinkAnnotation | PdfNoteAnnotation;

/** A marked-content sequence a page opened with `beginMarked`. */
export interface PdfMarkedContentRef {
	page: PdfVectorPage;
	mcid: number;
}

/** An annotation that belongs to a structure element (a Link's annotation). */
export interface PdfAnnotationRef {
	page: PdfVectorPage;
	annotation: PdfAnnotation;
}

export type PdfStructKid = PdfStructElement | PdfMarkedContentRef | PdfAnnotationRef;

/** One element of the logical structure (`Document`, `H1`, `P`, `Table`, `Figure`, `Link`...). */
export interface PdfStructElement {
	type: string;
	kids: PdfStructKid[];
	/** Alternate text (`/Alt`), for figures. */
	alt?: string;
	lang?: string;
	/**
	 * Standard structure attributes (`/A`, ISO 32000-1 14.8.5) by owner, each value a PDF name
	 * without its slash: `{ List: { ListNumbering: 'Decimal' } }`.
	 */
	attributes?: Readonly<Record<string, Readonly<Record<string, string>>>>;
}

/** One bookmark of the outline, with its nested bookmarks. */
export interface PdfOutlineItem {
	title: string;
	dest: PdfDestination;
	children?: PdfOutlineItem[];
}

export const num = (value: number): string => {
	if (!Number.isFinite(value)) return '0';
	const text = value.toFixed(4);
	return text.includes('.') ? text.replace(/\.?0+$/, '') || '0' : text;
};

/** A PDF text string: a literal for printable ASCII, else UTF-16BE hex with a byte order mark. */
export function pdfTextString(text: string): string {
	if (/^[\x20-\x7e]*$/.test(text)) return pdfLiteral(text);
	let hex = 'FEFF';
	for (let index = 0; index < text.length; index++)
		hex += text.charCodeAt(index).toString(16).padStart(4, '0').toUpperCase();
	return `<${hex}>`;
}

/** A PDF name for a structure type or tag, with characters outside the regular set escaped. */
export function pdfName(name: string): string {
	let out = '/';
	for (const byte of new TextEncoder().encode(name))
		out +=
			byte > 0x20 && byte < 0x7f && !'#()<>[]{}/%'.includes(String.fromCharCode(byte))
				? String.fromCharCode(byte)
				: `#${byte.toString(16).padStart(2, '0')}`;
	return out;
}

/** A URI as the 7-bit ASCII string a `/URI` action takes, with other characters escaped. */
const uriString = (uri: string): string =>
	pdfLiteral(uri.replace(/[^\x20-\x7e]+/g, (text) => encodeURIComponent(text)));

/** Writes indirect objects and allocates object numbers. */
export interface PdfObjectSink {
	object(id: number, body: string): void;
	/** Flate-compressed unless `raw`. */
	stream(id: number, dictionary: string, data: Uint8Array, raw?: boolean): void;
	allocate(): number;
	pageId(page: PdfVectorPage): number | undefined;
}

/** `[page /XYZ left top null]`, or undefined when the page is not in the file. */
export function destinationArray(dest: PdfDestination, sink: PdfObjectSink): string | undefined {
	const id = sink.pageId(dest.page);
	if (id === undefined) return undefined;
	return `[${id} 0 R /XYZ ${dest.x === undefined ? 'null' : num(dest.x)} ${num(dest.y)} null]`;
}

const rectArray = (rect: readonly number[]) => `[${rect.map(num).join(' ')}]`;

const NOTE_ICON =
	'1 0.84 0.2 rg 0.45 0.35 0 RG 1 w 0.5 0.5 19 17 re B 4 11 m 16 11 l 4 7 m 13 7 l S';

/**
 * Writes one annotation and returns its object number, or undefined when it has nowhere to go
 * (a link to a page outside the file). `structParent` keys it into the structure's parent tree.
 */
export function writeAnnotation(
	annotation: PdfAnnotation,
	sink: PdfObjectSink,
	structParent?: number,
): number | undefined {
	const parent = structParent === undefined ? '' : ` /StructParent ${structParent}`;
	if (annotation.kind === 'link') {
		let action: string | undefined;
		if (annotation.dest) {
			const dest = destinationArray(annotation.dest, sink);
			if (dest) action = `/Dest ${dest}`;
		} else if (annotation.uri) action = `/A << /S /URI /URI ${uriString(annotation.uri)} >>`;
		if (!action) return undefined;
		const id = sink.allocate();
		const contents = annotation.contents ? ` /Contents ${pdfTextString(annotation.contents)}` : '';
		sink.object(
			id,
			`<< /Type /Annot /Subtype /Link /Rect ${rectArray(annotation.rect)} /Border [0 0 0] /F 4 ${action}${contents}${parent} >>`,
		);
		return id;
	}
	// Notes carry their own appearance (a small icon), as PDF/A requires of non-link annotations.
	const appearance = sink.allocate();
	sink.stream(
		appearance,
		'/Type /XObject /Subtype /Form /BBox [0 0 20 18]',
		new TextEncoder().encode(NOTE_ICON),
	);
	const id = sink.allocate();
	const author = annotation.author ? ` /T ${pdfTextString(annotation.author)}` : '';
	const date = annotation.date ? ` /M ${pdfTextString(annotation.date)}` : '';
	sink.object(
		id,
		`<< /Type /Annot /Subtype /Text /Rect ${rectArray(annotation.rect)} /Contents ${pdfTextString(annotation.contents)}${author}${date} /Name /Comment /C [1 0.84 0.2] /F 28 /Open false /AP << /N ${appearance} 0 R >>${parent} >>`,
	);
	return id;
}

/** Writes the outline and returns the `/Outlines` dictionary's object number. */
export function writeOutline(items: readonly PdfOutlineItem[], sink: PdfObjectSink): number {
	const root = sink.allocate();
	const visible = (list: readonly PdfOutlineItem[]): number =>
		list.reduce((sum, item) => sum + 1 + visible(item.children ?? []), 0);
	const level = (list: readonly PdfOutlineItem[], parent: number): number[] => {
		const ids = list.map(() => sink.allocate());
		list.forEach((item, index) => {
			const id = ids[index]!;
			const kids = item.children?.length ? level(item.children, id) : [];
			const dest = destinationArray(item.dest, sink);
			const links = [
				index > 0 ? ` /Prev ${ids[index - 1]} 0 R` : '',
				index < ids.length - 1 ? ` /Next ${ids[index + 1]} 0 R` : '',
				kids.length
					? ` /First ${kids[0]} 0 R /Last ${kids.at(-1)} 0 R /Count ${visible(item.children!)}`
					: '',
			].join('');
			sink.object(
				id,
				`<< /Title ${pdfTextString(item.title)} /Parent ${parent} 0 R${links}${dest ? ` /Dest ${dest}` : ''} >>`,
			);
		});
		return ids;
	};
	const top = level(items, root);
	sink.object(
		root,
		top.length
			? `<< /Type /Outlines /First ${top[0]} 0 R /Last ${top.at(-1)} 0 R /Count ${visible(items)} >>`
			: '<< /Type /Outlines /Count 0 >>',
	);
	return root;
}

const isElement = (kid: PdfStructKid): kid is PdfStructElement => 'type' in kid;
const isAnnotationRef = (kid: PdfStructKid): kid is PdfAnnotationRef => 'annotation' in kid;

/** The annotations the structure tree refers to, which need a `/StructParent` key. */
export function structAnnotations(root: PdfStructElement): Set<PdfAnnotation> {
	const found = new Set<PdfAnnotation>();
	const walk = (element: PdfStructElement) => {
		for (const kid of element.kids)
			if (isElement(kid)) walk(kid);
			else if (isAnnotationRef(kid)) found.add(kid.annotation);
	};
	walk(root);
	return found;
}

/** The `/A` entry of an element: one attribute dictionary, or an array of them, by owner. */
function attributeObjects(attributes: PdfStructElement['attributes']): string {
	const dictionaries = Object.entries(attributes ?? {}).map(
		([owner, values]) =>
			`<< /O ${pdfName(owner)}${Object.entries(values)
				.map(([key, value]) => ` ${pdfName(key)} ${pdfName(value)}`)
				.join('')} >>`,
	);
	if (!dictionaries.length) return '';
	return dictionaries.length === 1 ? ` /A ${dictionaries[0]}` : ` /A [${dictionaries.join(' ')}]`;
}

/** Whether an element holds any content (marked content or an annotation) at any depth. */
const hasContent = (element: PdfStructElement): boolean =>
	element.kids.some((kid) => !isElement(kid) || hasContent(kid));

/**
 * Writes the structure tree and its parent tree. `pageKeys` gives each page's `/StructParents`
 * key; `annotations` each written annotation's object number and `/StructParent` key. Elements
 * with no content are left out. Returns the `/StructTreeRoot` object number.
 */
export function writeStructTree(
	root: PdfStructElement,
	sink: PdfObjectSink,
	pageKeys: ReadonlyMap<PdfVectorPage, number>,
	annotations: ReadonlyMap<PdfAnnotation, { id: number; key: number }>,
	nextKey: number,
): number {
	const treeId = sink.allocate();
	const ids = new Map<PdfStructElement, number>();
	const assign = (element: PdfStructElement) => {
		ids.set(element, sink.allocate());
		for (const kid of element.kids) if (isElement(kid) && hasContent(kid)) assign(kid);
	};
	assign(root);
	const pageParents = new Map<number, (number | undefined)[]>();
	const annotationParents = new Map<number, number>();
	const write = (element: PdfStructElement, parent: number) => {
		const id = ids.get(element)!;
		const kids: string[] = [];
		for (const kid of element.kids) {
			if (isElement(kid)) {
				if (ids.has(kid)) {
					write(kid, id);
					kids.push(`${ids.get(kid)} 0 R`);
				}
				continue;
			}
			const page = sink.pageId(kid.page);
			if (page === undefined) continue;
			if (isAnnotationRef(kid)) {
				const written = annotations.get(kid.annotation);
				if (!written) continue;
				annotationParents.set(written.key, id);
				kids.push(`<< /Type /OBJR /Pg ${page} 0 R /Obj ${written.id} 0 R >>`);
				continue;
			}
			const key = pageKeys.get(kid.page);
			if (key === undefined) continue;
			const parents = pageParents.get(key) ?? [];
			parents[kid.mcid] = id;
			pageParents.set(key, parents);
			kids.push(`<< /Type /MCR /Pg ${page} 0 R /MCID ${kid.mcid} >>`);
		}
		const extra = [
			element.alt !== undefined ? ` /Alt ${pdfTextString(element.alt)}` : '',
			element.lang ? ` /Lang ${pdfTextString(element.lang)}` : '',
			attributeObjects(element.attributes),
		].join('');
		sink.object(
			id,
			`<< /Type /StructElem /S ${pdfName(element.type)} /P ${parent} 0 R${extra} /K [${kids.join(' ')}] >>`,
		);
	};
	write(root, treeId);
	const nums: string[] = [];
	const keys = [...new Set([...pageParents.keys(), ...annotationParents.keys()])].sort(
		(a, b) => a - b,
	);
	for (const key of keys) {
		const parents = pageParents.get(key);
		if (parents)
			nums.push(
				`${key} [${Array.from(parents, (id) => (id === undefined ? 'null' : `${id} 0 R`)).join(' ')}]`,
			);
		else nums.push(`${key} ${annotationParents.get(key)} 0 R`);
	}
	const parentTree = sink.allocate();
	sink.object(parentTree, `<< /Nums [${nums.join(' ')}] >>`);
	sink.object(
		treeId,
		`<< /Type /StructTreeRoot /K [${ids.get(root)} 0 R] /ParentTree ${parentTree} 0 R /ParentTreeNextKey ${nextKey} >>`,
	);
	return treeId;
}
