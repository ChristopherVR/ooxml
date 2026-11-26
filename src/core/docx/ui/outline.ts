import { Fragment, type Node } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { defaultStyleCatalogs } from '../default-styles';
import { paragraphStyleId, styleCatalogsOf, type StyleCatalogs } from '../edit-styles';
import type { DocumentModel, Paragraph, ParagraphStyleCatalog } from '../model';
import { resolveParagraphFormatting } from '../paragraph-styles';
import { headingLevel } from '../table-of-contents';

export const OUTLINE_BODY = 10;

export interface OutlineEntry {
	index: number;
	id: string;
	pos: number;
	node: Node;
	level: number;
	indent: number;
	end: number;
	previous: number | undefined;
	section: number;
}

function levelOf(node: Node, catalog: ParagraphStyleCatalog | undefined): number {
	if (node.type.name !== 'paragraph') return OUTLINE_BODY;
	const paragraph: Paragraph = {
		type: 'paragraph',
		id: node.attrs.id,
		runs: [],
		...(node.attrs.style ? { style: node.attrs.style } : {}),
		...(node.attrs.outlineLevel != null ? { outlineLevel: node.attrs.outlineLevel } : {}),
	};
	const level =
		(catalog
			? resolveParagraphFormatting(paragraph, catalog).outlineLevel
			: paragraph.outlineLevel) ?? headingLevel(paragraph, catalog);
	return level && level >= 1 && level <= 9 ? level : OUTLINE_BODY;
}

/** Indexes top-level paragraphs and heading subtrees. Tables and other body blocks stay atomic. */
export function indexOutline(doc: Node, model: DocumentModel): OutlineEntry[] {
	const catalogs: StyleCatalogs | null | undefined = doc.attrs.styleCatalogs;
	const paragraphStyles = catalogs?.paragraphStyles ?? model.paragraphStyles;
	const entries: OutlineEntry[] = [];
	const parents: OutlineEntry[] = [];
	const ends = new Set(model.sections?.slice(0, -1).map((section) => section.endsAtBlockId));
	let section = 0;
	doc.forEach((node, pos, index) => {
		const level = levelOf(node, paragraphStyles);
		let closed: OutlineEntry | undefined;
		while (parents.length && (parents.at(-1)?.level ?? 0) >= level) {
			closed = parents.pop();
			if (closed) closed.end = index;
		}
		const preceding = level === OUTLINE_BODY ? entries.at(-1) : closed;
		const entry: OutlineEntry = {
			index,
			id: String(node.attrs.id ?? ''),
			pos,
			node,
			level,
			indent: level === OUTLINE_BODY ? (parents.at(-1)?.level ?? 0) : level - 1,
			end: level === OUTLINE_BODY ? index + 1 : doc.childCount,
			previous: preceding?.level === level ? preceding.index : undefined,
			section,
		};
		entries.push(entry);
		if (level < OUTLINE_BODY) parents.push(entry);
		if (ends.has(entry.id)) section++;
	});
	return entries;
}

/** A depth filter and collapsed headings affect presentation only, never the underlying blocks. */
export function visibleOutline(
	entries: readonly OutlineEntry[],
	showLevel: number,
	collapsed: ReadonlySet<string>,
): boolean[] {
	let hiddenUntil = 0;
	return entries.map((entry) => {
		if (entry.index < hiddenUntil || entry.level > showLevel) return false;
		if (entry.level < OUTLINE_BODY && collapsed.has(entry.id)) hiddenUntil = entry.end;
		return true;
	});
}

export function outlinePeer(
	entries: readonly OutlineEntry[],
	entry: OutlineEntry,
	direction: -1 | 1,
): OutlineEntry | undefined {
	const other = entries[direction === 1 ? entry.end : (entry.previous ?? -1)];
	if (!other || other.level !== entry.level) return undefined;
	const first = direction === 1 ? entry : other;
	const second = direction === 1 ? other : entry;
	// Never move a section terminator, or shift a heading across one.
	const after = entries[second.end];
	const last = entries[second.end - 1];
	if (
		!last ||
		first.end !== second.index ||
		first.section !== last.section ||
		(after && after.section !== first.section)
	)
		return undefined;
	return other;
}

/** Reorders whole outline branches while preserving their node IDs and the selection within them. */
export function moveOutline(
	state: EditorState,
	entries: readonly OutlineEntry[],
	entry: OutlineEntry,
	direction: -1 | 1,
): Transaction | undefined {
	const peer = outlinePeer(entries, entry, direction);
	if (!peer) return undefined;
	const first = direction === 1 ? entry : peer;
	const second = direction === 1 ? peer : entry;
	const end = entries[second.end]?.pos ?? state.doc.content.size;
	const selectionEnd = entries[entry.end]?.pos ?? state.doc.content.size;
	const { from, to, anchor, head } = state.selection;
	if (from < entry.pos || to >= selectionEnd) return undefined;
	const left = state.doc.slice(first.pos, second.pos).content;
	const right = state.doc.slice(second.pos, end).content;
	const shift = direction === 1 ? right.size : -left.size;
	const tr = state.tr.replaceWith(first.pos, end, Fragment.empty.append(right).append(left));
	return tr.setSelection(
		TextSelection.between(tr.doc.resolve(anchor + shift), tr.doc.resolve(head + shift)),
	);
}

export function outlineParagraphs(
	state: EditorState,
	entries: readonly OutlineEntry[],
): OutlineEntry[] {
	const { $from, $to } = state.selection;
	const first = $from.index(0);
	const end = $to.index(0) + ($to.depth > 0 ? 1 : 0);
	return entries
		.slice(first, Math.max(first + 1, end))
		.filter((entry) => entry.node.type.name === 'paragraph');
}

export function canChangeOutlineLevel(entry: OutlineEntry, direction: -1 | 1): boolean {
	return direction === 1 ? entry.level < OUTLINE_BODY : entry.level > 1;
}

function outlineStyle(level: number) {
	return level === OUTLINE_BODY
		? { id: 'BodyText', name: 'Body Text' }
		: { id: `Heading${level}`, name: `heading ${level}` };
}

function hasOutlineStyle(name: string, level: number): boolean {
	const style = outlineStyle(level);
	const key = name.toLowerCase();
	return key === style.name.toLowerCase() || key === style.id.toLowerCase();
}

/** Use Word's built-in styles for built-in headings and default body paragraphs. */
export function changeOutlineLevel(
	state: EditorState,
	model: DocumentModel,
	paragraphs: readonly OutlineEntry[],
	direction: -1 | 1,
): Transaction | undefined {
	return applyOutlineLevels(state, model, paragraphs, (entry) =>
		canChangeOutlineLevel(entry, direction) ? entry.level + direction : entry.level,
	);
}

export function setOutlineLevel(
	state: EditorState,
	model: DocumentModel,
	paragraphs: readonly OutlineEntry[],
	level: number,
): Transaction | undefined {
	if (!Number.isInteger(level) || level < 1 || level > OUTLINE_BODY) return undefined;
	return applyOutlineLevels(state, model, paragraphs, () => level);
}

function applyOutlineLevels(
	state: EditorState,
	model: DocumentModel,
	paragraphs: readonly OutlineEntry[],
	levelFor: (entry: OutlineEntry) => number,
): Transaction | undefined {
	const tr = state.tr;
	const catalogs = styleCatalogsOf(model);
	const catalog = catalogs.paragraphStyles ?? { styles: {}, warnings: [], docDefaults: {} };
	const styles = { ...catalog.styles };
	let stylesChanged = false;
	for (const entry of paragraphs) {
		const level = levelFor(entry);
		if (level === entry.level) continue;
		const styleId = typeof entry.node.attrs.style === 'string' ? entry.node.attrs.style : '';
		const name = styles[styleId]?.name ?? styleId;
		const defaultBody =
			entry.level === OUTLINE_BODY &&
			(!styleId || styleId === paragraphStyleId(catalogs, undefined));
		if (
			entry.node.attrs.outlineLevel != null ||
			(!defaultBody && !hasOutlineStyle(name, entry.level))
		) {
			tr.setNodeAttribute(entry.pos, 'outlineLevel', level === OUTLINE_BODY ? 0 : level);
			continue;
		}
		const existing = Object.values(styles).find((style) =>
			hasOutlineStyle(style.name ?? style.id, level),
		);
		if (existing) {
			tr.setNodeAttribute(entry.pos, 'style', existing.id);
			continue;
		}
		const target = outlineStyle(level);
		let id = target.id;
		for (
			let suffix = 2;
			Object.hasOwn(styles, id) || Object.hasOwn(catalogs.characterStyles?.styles ?? {}, id);
			suffix++
		)
			id = target.id + suffix;
		const template = defaultStyleCatalogs().paragraphStyles.styles[target.id];
		const base = paragraphStyleId(catalogs, undefined);
		styles[id] = {
			id,
			name: target.name,
			...(base && { basedOn: base, next: base }),
			formatting: { ...template?.formatting, outlineLevel: level === OUTLINE_BODY ? 0 : level },
		};
		stylesChanged = true;
		tr.setNodeAttribute(entry.pos, 'style', id);
	}
	if (stylesChanged)
		tr.setDocAttribute('styleCatalogs', { ...catalogs, paragraphStyles: { ...catalog, styles } });
	return tr.docChanged ? tr : undefined;
}
