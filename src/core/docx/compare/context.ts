import type { DocumentModel, Paragraph, Revision, TableCell, TextRun } from '../model';
import type { CompareGranularity, CompareSummary } from './types';

export type ChangeKind = 'insert' | 'delete' | 'formatChange';

/** Shared state of one comparison: options, the revision factory and what it reports. */
export interface CompareContext {
	granularity: CompareGranularity;
	/** Revised paragraphs never aligned with the original: they show inserted with their runs kept. */
	whole: ReadonlySet<string>;
	compareFormatting: boolean;
	/** A new revision with a provisional id; saving numbers it with the document's allocator. */
	revision(kind: ChangeKind, previousRunPropertiesXml?: string): Revision;
	/** A new block id, for content carried over from the original. */
	blockId(): string;
	/** A run of the original as deleted text in the revised package, or undefined to drop it. */
	importRun(run: TextRun): TextRun | undefined;
	/** The properties of an original paragraph whose deleted mark the result shows. */
	importParagraph(paragraph: Paragraph): Paragraph;
	/** The properties of an original table cell shown in a deleted row. */
	importCell(cell: TableCell): TableCell;
	/** The paragraph-id layout of a table in the revised package as saved, before acceptance. */
	sourceLayout(tableId: string): string | undefined;
	/** The style id when the revised package defines it, else undefined. */
	tableStyle(style: string | undefined): string | undefined;
	summary: CompareSummary;
	warn(message: string): void;
}

export interface ContextOptions {
	author: string;
	date: string;
	granularity: CompareGranularity;
	compareFormatting: boolean;
	/** The revised document: its catalogs decide which original styles and lists can be kept. */
	revised: DocumentModel;
	/** Table layouts of the revised package as loaded (`tableLayout`), keyed by table id. */
	sourceLayouts: ReadonlyMap<string, string>;
	whole: ReadonlySet<string>;
	warnings: string[];
}

/** Word's revision date form: seconds, no fraction (`2026-11-15T07:26:00Z`). */
export function revisionDate(date: string | Date | undefined): string {
	const value = date === undefined ? new Date() : typeof date === 'string' ? new Date(date) : date;
	if (Number.isNaN(value.getTime())) throw new Error(`Invalid revision date: ${String(date)}`);
	return value.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

export function createContext(options: ContextOptions): CompareContext {
	let revisions = 0;
	let blocks = 0;
	const { revised } = options;
	const warnings = new Set<string>();
	const warn = (message: string) => {
		if (warnings.has(message)) return;
		warnings.add(message);
		options.warnings.push(message);
	};
	const characterStyles = revised.characterStyles?.styles ?? {};
	const paragraphStyles = revised.paragraphStyles?.styles ?? {};
	const tableStyles = revised.tableStyles?.styles ?? {};
	const nums = revised.numberingCatalog?.nums ?? {};
	const summary: CompareSummary = { insertions: 0, deletions: 0, formatChanges: 0 };
	return {
		granularity: options.granularity,
		whole: options.whole,
		compareFormatting: options.compareFormatting,
		revision(kind, previousRunPropertiesXml) {
			if (kind === 'insert') summary.insertions++;
			else if (kind === 'delete') summary.deletions++;
			else summary.formatChanges++;
			const revision: Revision = {
				kind,
				author: options.author,
				date: options.date,
				id: `compare-${++revisions}`,
			};
			if (previousRunPropertiesXml !== undefined)
				revision.previousRunPropertiesXml = previousRunPropertiesXml;
			return revision;
		},
		blockId: () => `compare-block-${++blocks}`,
		importRun(source) {
			if (source.image) {
				warn('Pictures and drawings deleted from the original are not shown in the comparison.');
				return undefined;
			}
			if (source.noteReference) {
				warn('Footnote and endnote references deleted from the original are not shown.');
				return undefined;
			}
			const run: TextRun = { ...source };
			delete run.revision;
			delete run.formatRevision;
			delete run.commentIds;
			delete run.contentControl;
			delete run.outerContentControls;
			delete run.restoredRunPropertiesXml;
			if (run.style && !characterStyles[run.style]) delete run.style;
			return run;
		},
		importParagraph(source) {
			const paragraph: Paragraph = { ...source, runs: [] };
			for (const key of [
				'markRevision',
				'formatRevision',
				'contentControl',
				'bookmarks',
				'permissions',
				'sourceParagraphPropertiesXml',
				'restoredParagraphPropertiesXml',
				'dropCap',
			] as const)
				delete paragraph[key];
			if (paragraph.style && !paragraphStyles[paragraph.style]) delete paragraph.style;
			if (paragraph.numbering && !nums[String(paragraph.numbering.numId)])
				delete paragraph.numbering;
			return paragraph;
		},
		importCell(source) {
			const cell: TableCell = { ...source, paragraphs: [] };
			if (cell.nestedTables?.length || cell.sourceBlocks)
				warn('Tables nested in deleted table rows are not shown in the comparison.');
			delete cell.nestedTables;
			delete cell.sourceBlocks;
			delete cell.contentControl;
			return cell;
		},
		sourceLayout: (tableId) => options.sourceLayouts.get(tableId),
		tableStyle: (style) => (style && tableStyles[style] ? style : undefined),
		summary,
		warn,
	};
}
