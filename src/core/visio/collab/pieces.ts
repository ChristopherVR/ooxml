// Splits one `Shape` element into the fine-grained pieces the shared drawing keeps one entry for,
// and builds it back. Pieces are addressed by a path inside the shape, so two peers that edit
// different cells of one shape touch different entries:
//   c/<N>                       a top-level Cell
//   s/<N[:IX]>                  a Section (its opening tag)
//   s/<sec>/c/<N>               a Cell of the section
//   s/<sec>/r/<IX|N|#i>         a Row (its opening tag)
//   s/<sec>/r/<row>/c/<N>       a Cell of the row
//   x/<Name>#<k>                any other child (Text, Data, ForeignData, ...) as a whole
// Nested shapes are not pieces: each one is its own shape entry that names its parent.
// The XML of a piece is the text the editor wrote, so cells round-trip exactly; whitespace between
// elements is not kept.
import { canon } from './package';
import { attrOf, scanContainer, type ElementSpan } from './scan';

export interface Piece {
	path: string;
	/** Path of the container holding it: '' for the shape itself. */
	parent: string;
	/** A container (section, row) keeps its opening tag; a leaf its whole element. */
	kind: 'open' | 'leaf';
	text: string;
	/** See {@link PieceEntry.r}; set by the snapshot, which knows the page's ids and keys. */
	refs?: string[];
}

export interface SplitShape {
	/** The opening tag, never self-closing. */
	open: string;
	pieces: Piece[];
	/** The nested `Shape` elements, in order. */
	nested: string[];
}

const seg = (value: string): string => value.replace(/[%/]/g, (c) => (c === '%' ? '%25' : '%2F'));
const tagName = (open: string): string => /^<([^\s/>]+)/.exec(open)?.[1] ?? 'Shape';

/** An opening tag as text, made non-self-closing. */
export const openTagOf = (text: string, span: ElementSpan): string => {
	const tag = text.slice(span.start, span.openEnd);
	return span.selfClosing ? tag.replace(/\s*\/>$/, '>') : tag;
};
export const closeTagOf = (open: string): string => `</${tagName(open)}>`;

/** Splits the element `xml` (a `Shape`) into its opening tag, pieces and nested shapes. */
export function splitShape(xml: string): SplitShape | undefined {
	const box = scanContainer(xml, 'Shape');
	if (!box) return undefined;
	const used = new Set<string>();
	const unique = (path: string): string => {
		let out = path;
		for (let n = 2; used.has(out); n++) out = `${path}~${n}`;
		used.add(out);
		return out;
	};
	const pieces: Piece[] = [];
	const nested: string[] = [];
	const others = new Map<string, number>();
	const other = (parent: string, name: string, text: string): void => {
		const k = others.get(parent + name) ?? 0;
		others.set(parent + name, k + 1);
		pieces.push({
			path: unique(`${parent ? `${parent}/` : ''}x/${name}#${k}`),
			parent,
			kind: 'leaf',
			text,
		});
	};
	const cell = (parent: string, text: string): void => {
		const n = attrOf(text.slice(0, text.search(/>/) + 1), 'N');
		if (n === undefined) return other(parent, 'Cell', text);
		pieces.push({
			path: unique(`${parent ? `${parent}/` : ''}c/${seg(n)}`),
			parent,
			kind: 'leaf',
			text,
		});
	};
	for (const child of box.children) {
		const text = xml.slice(child.start, child.end);
		if (child.name === 'Cell') cell('', text);
		else if (child.name === 'Shapes') {
			const list = scanContainer(text, 'Shapes');
			for (const s of list?.children ?? [])
				if (s.name === 'Shape') nested.push(text.slice(s.start, s.end));
		} else if (child.name === 'Section') {
			const head = openTagOf(xml, child);
			const n = attrOf(head, 'N');
			const ix = attrOf(head, 'IX');
			if (n === undefined && ix === undefined) {
				other('', 'Section', text);
				continue;
			}
			const sec = unique(`s/${seg(n ?? '')}${ix !== undefined ? `:${seg(ix)}` : ''}`);
			pieces.push({ path: sec, parent: '', kind: 'open', text: head });
			const inner = scanContainer(text, 'Section');
			let rowIndex = 0;
			for (const part of inner?.children ?? []) {
				const sub = text.slice(part.start, part.end);
				if (part.name === 'Cell') cell(sec, sub);
				else if (part.name === 'Row') {
					const rowHead = openTagOf(text, part);
					const id = attrOf(rowHead, 'IX') ?? attrOf(rowHead, 'N') ?? `#${rowIndex}`;
					rowIndex++;
					const row = unique(`${sec}/r/${seg(id)}`);
					pieces.push({ path: row, parent: sec, kind: 'open', text: rowHead });
					for (const rc of scanContainer(sub, 'Row')?.children ?? [])
						if (rc.name === 'Cell') cell(row, sub.slice(rc.start, rc.end));
						else other(row, rc.name, sub.slice(rc.start, rc.end));
				} else other(sec, part.name, sub);
			}
		} else other('', child.name, text);
	}
	return { open: openTagOf(xml, box), pieces, nested };
}

export interface PieceEntry {
	pos: number;
	/** Whole element of a leaf. */
	xml?: string;
	/** Opening tag of a container. */
	open?: string;
	/**
	 * For a cell whose formula names shapes (`Sheet.<id>`): pairs of referenced id and the key of the
	 * shape that id meant to the author, flattened (`[id, key, id, key]`). Not part of the content.
	 */
	r?: string[];
}

/** Builds the element: `open`, the pieces under `parent` in position order, then `tail`. */
export function buildShape(
	open: string,
	entries: ReadonlyMap<string, PieceEntry>,
	tail: string,
): string {
	const byParent = new Map<string, { path: string; entry: PieceEntry }[]>();
	for (const [path, entry] of entries) {
		const parent = parentOf(path);
		const list = byParent.get(parent) ?? [];
		list.push({ path, entry });
		byParent.set(parent, list);
	}
	const render = (parent: string): string =>
		(byParent.get(parent) ?? [])
			.sort((a, b) => a.entry.pos - b.entry.pos || (a.path < b.path ? -1 : a.path > b.path ? 1 : 0))
			.map(({ path, entry }) =>
				entry.open !== undefined
					? entry.open + render(path) + closeTagOf(entry.open)
					: (entry.xml ?? ''),
			)
			.join('');
	return open + render('') + tail + closeTagOf(open);
}

/** The container path of a piece: its path without the last `<kind>/<name>` pair. */
export function parentOf(path: string): string {
	const parts = path.split('/');
	return parts.length <= 2 ? '' : parts.slice(0, -2).join('/');
}

/** A form of a piece that ignores how the writer serialised it (quote style, empty elements). */
const SIMPLE =
	/^<([^\s/>]+)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(?:\/>|>([^<]*)<\/\1>)$/;
const ATTR = /\s([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
export function pieceContent(kind: 'open' | 'leaf', text: string): string {
	const simple = SIMPLE.exec(kind === 'open' ? `${text}</${tagName(text)}>` : text);
	if (!simple) return kind === 'open' ? canon(`${text}</${tagName(text)}>`) : canon(text);
	const attrs = [...(simple[2] ?? '').matchAll(ATTR)]
		.map((m) => [m[1]!, m[2] ?? m[3] ?? ''] as const)
		.sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
	return JSON.stringify([simple[1], attrs, simple[3] ?? '']);
}
