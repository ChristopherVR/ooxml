/**
 * table-cell-source.ts: pair a loaded table cell's runs with the markup they
 * were read from, so an edit can keep each run's own `a:rPr` (hyperlinks,
 * theme colours and their transforms, fields, language, effects).
 *
 * The cell's `textRuns` come from `extractTableCellTextRuns`, which walks
 * every `a:p` and its `a:r` / `a:br` / `a:fld` children in authored order;
 * this walks the same children the same way and checks each pairing by tag
 * and text, so a cell whose markup no longer matches its runs gets no source
 * at all rather than a wrong one.
 *
 * @module render/table-cell-source
 */
import type {
	PptxTableCellTextRun,
	TableCellRunSource,
	TablePptxElement,
	TextStyle,
	XmlObject,
} from 'ooxml-core/pptx';

import { paragraphContentEntries } from '../../core/core/runtime/paragraph-sibling-order';
import { ensureArray, getTblFromRawXml } from '../../core/core/runtime/table-structural-helpers';
import { xmlText } from '../../core/utils';

const CELL_CONTENT_TAGS: ReadonlySet<string> = new Set(['a:r', 'a:br', 'a:fld']);

/** One cell's source markup, aligned with its runs. */
export interface TableCellSource {
	/** The source `a:p` of each paragraph, in order. */
	paragraphs: XmlObject[];
	/** For each entry of `textRuns`, its source run (`undefined` for a paragraph break). */
	runs: Array<Omit<TableCellRunSource, 'style'> | undefined>;
}

/** The `a:tc` of grid position `(row, col)` in an element's markup. */
export function tableCellSourceXml(
	element: TablePptxElement,
	row: number,
	col: number,
): XmlObject | undefined {
	const tbl = element.rawXml ? getTblFromRawXml(element.rawXml) : undefined;
	const rows = ensureArray(tbl?.['a:tr'] as XmlObject | XmlObject[] | undefined);
	const cells = ensureArray(rows[row]?.['a:tc'] as XmlObject | XmlObject[] | undefined);
	return cells[col];
}

/**
 * Align `runs` with the `a:txBody` of `cellXml`, or `undefined` when they do
 * not describe the same content.
 */
export function pairTableCellSource(
	runs: readonly PptxTableCellTextRun[],
	cellXml: XmlObject | undefined,
): TableCellSource | undefined {
	const txBody = cellXml?.['a:txBody'];
	if (!txBody || typeof txBody !== 'object') {
		return undefined;
	}
	const paragraphs = ensureArray((txBody as XmlObject)['a:p'] as XmlObject | XmlObject[]).map(
		(paragraph) => (paragraph && typeof paragraph === 'object' ? paragraph : ({} as XmlObject)),
	);
	const entries = paragraphs.map(
		(paragraph) => paragraphContentEntries(paragraph, CELL_CONTENT_TAGS).entries,
	);
	const paired: TableCellSource['runs'] = [];
	let paragraph = 0;
	let entry = 0;
	for (const run of runs) {
		if (run.isParagraphBreak) {
			if (entry !== (entries[paragraph]?.length ?? 0)) {
				return undefined;
			}
			paragraph++;
			entry = 0;
			paired.push(undefined);
			continue;
		}
		const next = entries[paragraph]?.[entry++];
		if (!next) {
			return undefined;
		}
		const [tag, item] = next;
		const node = item && typeof item === 'object' ? (item as XmlObject) : ({} as XmlObject);
		const expected = run.isLineBreak ? 'a:br' : run.isField ? 'a:fld' : 'a:r';
		if (tag !== expected || (tag !== 'a:br' && (xmlText(node['a:t']) ?? '') !== run.text)) {
			return undefined;
		}
		paired.push({ tag: expected, xml: node });
	}
	if (paragraph !== paragraphs.length - 1 || entry !== (entries[paragraph]?.length ?? 0)) {
		return undefined;
	}
	return { paragraphs, runs: paired };
}

/** A run source with the style the run was read as. */
export function withRunStyle(
	source: Omit<TableCellRunSource, 'style'> | undefined,
	style: TextStyle,
): TableCellRunSource | undefined {
	return source ? { ...source, style: { ...style } } : undefined;
}
