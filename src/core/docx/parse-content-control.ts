// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Where content controls (`w:sdt`) sit in a story: around runs, around blocks, rows and cells.
import type { ContentControl } from './content-control-model';
import { isSdt, parseContentControl, sdtContent } from './content-control';
import type { Revision, TextRun } from './model';
import { collectParagraphRuns, type OpenMoves } from './parse-revisions';
import { isElement, named, type XmlElement } from './xml';

/** The runs of an inline control, each tagged with it (an empty control keeps one empty run). */
export function parseInlineControl(
	sdt: XmlElement,
	parseRun: (run: XmlElement, revision?: Revision) => TextRun,
	parseOther: (item: XmlElement) => TextRun[] | undefined,
	resolveLink: (hyperlink: XmlElement) => TextRun['link'],
	comments: string[] | undefined,
	moves: OpenMoves | undefined,
): TextRun[] {
	const control = parseContentControl(sdt, 'inline');
	const content = sdtContent(sdt);
	const runs = content
		? collectParagraphRuns(content, parseRun, parseOther, resolveLink, comments, moves).runs
		: [];
	if (!runs.length) return [{ text: '', contentControl: control }];
	for (const run of runs) run.contentControl ??= control;
	return runs;
}

/** A block-level element found through any number of `w:sdt` wrappers. */
export interface FlatBlock {
	element: XmlElement;
	/** The innermost block-level control around it. */
	control?: ContentControl;
	/** The innermost `w:sdt` element around it. */
	sdt?: XmlElement;
}

/** The `local` children of `parent`, looking through `w:sdt/w:sdtContent` wrappers in order. */
export function flatChildren(
	parent: XmlElement,
	names: string[],
	level: 'block' | 'row' | 'cell' = 'block',
	enclosing?: { control: ContentControl; sdt: XmlElement },
	controls = new Map<XmlElement, ContentControl>(),
): FlatBlock[] {
	const found: FlatBlock[] = [];
	for (const child of Array.from(parent.childNodes).filter(isElement)) {
		if (isSdt(child)) {
			const content = sdtContent(child);
			if (!content) continue;
			let control = controls.get(child);
			if (!control) {
				control = parseContentControl(child, level);
				controls.set(child, control);
			}
			found.push(...flatChildren(content, names, level, { control, sdt: child }, controls));
		} else if (names.some((name) => named(child, name)))
			found.push({
				element: child,
				...(enclosing ? { control: enclosing.control, sdt: enclosing.sdt } : {}),
			});
	}
	return found;
}

/** Paragraphs and tables of a story container, through block-level content controls. */
export const flatBlocks = (container: XmlElement): FlatBlock[] =>
	flatChildren(container, ['p', 'tbl']);
/** Rows of a table, through row-level content controls. */
export const flatRows = (table: XmlElement): FlatBlock[] => flatChildren(table, ['tr'], 'row');
/** Cells of a row, through cell-level content controls. */
export const flatCells = (row: XmlElement): FlatBlock[] => flatChildren(row, ['tc'], 'cell');
/** Paragraphs of a cell, through block-level content controls. */
export const flatCellParagraphs = (cell: XmlElement): FlatBlock[] => flatChildren(cell, ['p']);
