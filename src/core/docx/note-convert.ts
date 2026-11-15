// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word's Convert Notes (References > Footnotes dialog > Convert, and "Convert to Endnote" /
// "Convert to Footnote" on a note): moves note content between word/footnotes.xml and
// word/endnotes.xml and retargets the reference runs. Word renumbers the displayed numbers by
// document order, which numberNotesInOrder already does; ids stay unique per part.
import type { Block, DocumentModel, Note, Paragraph, TextRun } from './model';

export type NoteKind = 'footnote' | 'endnote';

/** The three choices of Word's Convert Notes dialog. */
export type NoteConversion = 'footnotesToEndnotes' | 'endnotesToFootnotes' | 'swap';

/** Which notes to move: every note of a kind, or one note. */
export type NoteSelection = NoteConversion | { kind: NoteKind; id: string };

type NoteParts = Pick<DocumentModel, 'footnotes' | 'endnotes'>;
type StyleCatalogs = Pick<DocumentModel, 'paragraphStyles' | 'characterStyles'>;

export interface NoteConversionPlan {
	footnotes: Note[];
	endnotes: Note[];
	/** `kind:id` of each moved note to its new kind and id. */
	moved: Map<string, { kind: NoteKind; id: string }>;
	/** Notes left where they were because they hold content that cannot be moved yet. */
	skipped: { kind: NoteKind; id: string }[];
}

const KEY = { footnote: 'footnotes', endnote: 'endnotes' } as const;
const PREFIX = { footnote: 'fn', endnote: 'en' } as const;
const OTHER: Record<NoteKind, NoteKind> = { footnote: 'endnote', endnote: 'footnote' };
const TEXT_STYLE = { footnote: 'FootnoteText', endnote: 'EndnoteText' } as const;
const REFERENCE_STYLE = { footnote: 'FootnoteReference', endnote: 'EndnoteReference' } as const;

/** Maps every paragraph of a block list, including table cells and nested tables. */
function mapParagraphs(blocks: Block[], map: (paragraph: Paragraph) => Paragraph): Block[] {
	return blocks.map((block) =>
		block.type === 'paragraph'
			? map(block)
			: {
					...block,
					rows: block.rows.map((row) =>
						row.map((cell) => ({
							...cell,
							paragraphs: cell.paragraphs.map(map),
							...(cell.sourceBlocks ? { sourceBlocks: mapParagraphs(cell.sourceBlocks, map) } : {}),
						})),
					),
				},
	);
}

function eachRun(blocks: Block[], visit: (run: TextRun) => void): void {
	mapParagraphs(blocks, (paragraph) => {
		paragraph.runs.forEach(visit);
		return paragraph;
	});
}

/**
 * Pictures are rewritten from the model in the other part; drawings kept only as source XML
 * (charts, SmartArt, shapes, text boxes, floating pictures) cannot follow yet.
 */
function canMove(note: Note): boolean {
	let movable = true;
	eachRun(note.blocks, (run) => {
		const image = run.image;
		if (image && (!image.partName || image.unsupported || image.anchored || image.chart))
			movable = false;
	});
	return movable;
}

/**
 * The style a note's text or reference takes in the other kind: Word swaps FootnoteText and
 * FootnoteReference for their Endnote counterparts. When the document does not define the
 * counterpart, the original style stays so the note keeps its look.
 */
function mappedStyle(
	style: string | undefined,
	from: NoteKind,
	names: typeof TEXT_STYLE | typeof REFERENCE_STYLE,
	defined: Record<string, unknown> | undefined,
): string | undefined {
	if (style !== names[from]) return style;
	const target = names[OTHER[from]];
	return defined?.[target] ? target : style;
}

/** A reference run or note number mark moved to the other kind; other formatting is kept. */
export function convertNoteRun(run: TextRun, from: NoteKind, styles: StyleCatalogs): TextRun {
	const style = mappedStyle(run.style, from, REFERENCE_STYLE, styles.characterStyles?.styles);
	return style === run.style ? { ...run } : { ...run, style: style! };
}

function convertNoteBlocks(note: Note, from: NoteKind, id: string, styles: StyleCatalogs): Block[] {
	const to = OTHER[from];
	const oldPrefix = `${PREFIX[from]}${note.id}-`;
	const newPrefix = `${PREFIX[to]}${id}-`;
	return mapParagraphs(structuredClone(note.blocks), (paragraph) => {
		const style = mappedStyle(paragraph.style, from, TEXT_STYLE, styles.paragraphStyles?.styles);
		return {
			...paragraph,
			...(style !== undefined ? { style } : {}),
			id: paragraph.id.startsWith(oldPrefix)
				? newPrefix + paragraph.id.slice(oldPrefix.length)
				: paragraph.id,
			runs: paragraph.runs.map((run) =>
				run.noteMark === from ? { ...convertNoteRun(run, from, styles), noteMark: to } : run,
			),
		};
	});
}

function selects(selection: NoteSelection, kind: NoteKind, id: string): boolean {
	if (typeof selection === 'object') return selection.kind === kind && selection.id === id;
	if (selection === 'swap') return true;
	return selection === (kind === 'footnote' ? 'footnotesToEndnotes' : 'endnotesToFootnotes');
}

/** The highest id in use, never below 1 so new notes clear both separator id conventions. */
function maxId(notes: readonly Note[] | undefined, reserved: Iterable<string> = []): number {
	const ids = [...(notes ?? []).map((note) => note.id), ...reserved].map(Number);
	return Math.max(1, ...ids.filter(Number.isFinite));
}

/**
 * Plans moving the selected notes to the other part. `reservedIds` holds ids each part already
 * uses on disk (such as notes deleted this session), so a moved note never takes one over.
 */
export function planNoteConversion(
	parts: NoteParts,
	selection: NoteSelection,
	styles: StyleCatalogs = {},
	reservedIds: Partial<Record<NoteKind, Iterable<string>>> = {},
): NoteConversionPlan {
	const plan: NoteConversionPlan = {
		footnotes: [],
		endnotes: [],
		moved: new Map(),
		skipped: [],
	};
	const next = {
		footnote: maxId(parts.footnotes, reservedIds.footnote),
		endnote: maxId(parts.endnotes, reservedIds.endnote),
	};
	const incoming: Record<NoteKind, Note[]> = { footnote: [], endnote: [] };
	for (const from of ['footnote', 'endnote'] as const) {
		for (const note of parts[KEY[from]] ?? []) {
			if (!selects(selection, from, note.id)) {
				plan[KEY[from]].push(note);
				continue;
			}
			if (!canMove(note)) {
				plan.skipped.push({ kind: from, id: note.id });
				plan[KEY[from]].push(note);
				continue;
			}
			const to = OTHER[from];
			const id = String(++next[to]);
			plan.moved.set(`${from}:${note.id}`, { kind: to, id });
			incoming[to].push({ id, blocks: convertNoteBlocks(note, from, id, styles) });
		}
	}
	plan.footnotes.push(...incoming.footnote);
	plan.endnotes.push(...incoming.endnote);
	return plan;
}

/** The reference run retargeted by a plan, or the same run when its note did not move. */
export function retargetNoteReference(
	run: TextRun,
	plan: NoteConversionPlan,
	styles: StyleCatalogs,
): TextRun {
	const reference = run.noteReference;
	const target = reference && plan.moved.get(`${reference.kind}:${reference.id}`);
	if (!reference || !target) return run;
	return { ...convertNoteRun(run, reference.kind, styles), noteReference: { ...target } };
}

/**
 * Word's Convert Notes on a whole model: moves the selected notes to the other part and
 * retargets their references in the body. Emptied lists stay as `[]` so saving removes the
 * notes from their old part. Notes the plan skipped are reported in `skipped`.
 */
export function convertNotes(
	model: DocumentModel,
	selection: NoteSelection,
	reservedIds?: Partial<Record<NoteKind, Iterable<string>>>,
): { model: DocumentModel; plan: NoteConversionPlan } {
	const plan = planNoteConversion(model, selection, model, reservedIds);
	if (!plan.moved.size) return { model, plan };
	const blocks = mapParagraphs(model.blocks, (paragraph) =>
		paragraph.runs.some((run) => run.noteReference)
			? { ...paragraph, runs: paragraph.runs.map((run) => retargetNoteReference(run, plan, model)) }
			: paragraph,
	);
	return {
		model: { ...model, blocks, footnotes: plan.footnotes, endnotes: plan.endnotes },
		plan,
	};
}
