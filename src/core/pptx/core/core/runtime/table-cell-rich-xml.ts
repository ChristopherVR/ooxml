/**
 * table-cell-rich-xml.ts: write an edited table cell's paragraphs from its
 * segments while keeping the markup each run was read from.
 *
 * A segment that came from a source run (`TextSegment.tableRun`) re-emits
 * that run's own `a:rPr` verbatim (hyperlinks with their relationship ids,
 * scheme colours with their transforms, language, effects, anything this
 * model does not type) and overrides only the fields whose style changed
 * since it was read. A field (`a:fld`) is written back as the field while
 * its text is unchanged; once its text is edited it becomes a plain run, as
 * in PowerPoint. Each paragraph keeps its source `a:pPr` and `a:endParaRPr`.
 * Segments with no source (text typed into an empty cell) get run
 * properties from their style through `runPropertiesFromStyle`.
 *
 * @module runtime/table-cell-rich-xml
 */
import type { TableCellRunSource, TextSegment, TextStyle, XmlObject } from '../../types';
import { xmlText } from '../../utils';
import { writeParagraphChildren } from './paragraph-child-assembly';
import type { ParagraphChild } from './paragraph-child-assembly';
import { applyCellRunStyleUpdates } from './table-cell-run-style-xml';

type RunPropertiesFromStyle = (style: TextStyle) => XmlObject | undefined;

const PT_PER_PX = 72 / 96;

const DIFF_KEYS = [
	'bold',
	'italic',
	'underline',
	'strikethrough',
	'color',
	'fontFamily',
	'fontSize',
	'baseline',
	'characterSpacing',
] as const;

/** Whether any segment carries source markup, so the cell takes this writer. */
export function hasTableCellSource(segments: readonly TextSegment[]): boolean {
	return segments.some((segment) => segment.tableRun || segment.tableParagraph);
}

function clone<T>(value: T): T {
	return value === undefined ? value : structuredClone(value);
}

/** A source `a:rPr` (a bare `<a:rPr/>` parses as `''`) as an object to write into. */
function sourceRunProperties(source: TableCellRunSource): XmlObject | undefined {
	const rPr = source.xml['a:rPr'];
	if (rPr === undefined) {
		return undefined;
	}
	return rPr && typeof rPr === 'object' ? (clone(rPr) as XmlObject) : {};
}

/** The source run properties with every changed style field written over them. */
function runProperties(segment: TextSegment, source: TableCellRunSource): XmlObject | undefined {
	const rPr = sourceRunProperties(source);
	const before = source.style;
	const after = segment.style ?? {};
	const updates: Record<string, unknown> = {};
	for (const key of DIFF_KEYS) {
		if (before[key] !== after[key]) {
			updates[key] = after[key];
		}
	}
	if (Object.keys(updates).length === 0) {
		return rPr;
	}
	const target = rPr ?? {};
	if (typeof updates.fontSize === 'number') {
		updates.fontSize = Math.round(updates.fontSize * PT_PER_PX * 100) / 100;
	}
	applyCellRunStyleUpdates(target, updates);
	// Turned off over an inherited (table style) format: say so explicitly.
	const off: Record<string, [string, string]> = {
		bold: ['@_b', '0'],
		italic: ['@_i', '0'],
		underline: ['@_u', 'none'],
		strikethrough: ['@_strike', 'noStrike'],
	};
	for (const [key, [attribute, value]] of Object.entries(off)) {
		if (key in updates && updates[key] === false) {
			target[attribute] = value;
		}
	}
	if ('baseline' in updates) {
		if (updates.baseline) {
			target['@_baseline'] = String(updates.baseline);
		} else {
			delete target['@_baseline'];
		}
	}
	if ('characterSpacing' in updates) {
		if (updates.characterSpacing) {
			target['@_spc'] = String(updates.characterSpacing);
		} else {
			delete target['@_spc'];
		}
	}
	return target;
}

function textRun(rPr: XmlObject | undefined, text: string): XmlObject {
	const run: XmlObject = {};
	if (rPr !== undefined) {
		run['a:rPr'] = rPr;
	}
	run['a:t'] = text;
	return run;
}

/** The `a:p` child for one segment, or `undefined` for empty text. */
function segmentChild(
	segment: TextSegment,
	runPropertiesFromStyle: RunPropertiesFromStyle,
): ParagraphChild | undefined {
	const source = segment.tableRun;
	if (segment.isLineBreak) {
		const br = source?.tag === 'a:br' && source.xml ? clone(source.xml) : {};
		return { tag: 'a:br', value: br && typeof br === 'object' ? br : {} };
	}
	if (!segment.text) {
		return undefined;
	}
	if (!source || source.tag === 'a:br') {
		return {
			tag: 'a:r',
			value: textRun(runPropertiesFromStyle(segment.style ?? {}), segment.text),
		};
	}
	const rPr = runProperties(segment, source);
	const unchangedField =
		source.tag === 'a:fld' &&
		segment.fieldType !== undefined &&
		segment.text === (xmlText(source.xml['a:t']) ?? '');
	if (!unchangedField) {
		return { tag: 'a:r', value: textRun(rPr, segment.text) };
	}
	// CT_TextField: attributes, then rPr?, pPr?, t?.
	const field: XmlObject = {};
	for (const [key, value] of Object.entries(source.xml)) {
		if (key.startsWith('@_')) {
			field[key] = value;
		}
	}
	if (rPr !== undefined) {
		field['a:rPr'] = rPr;
	}
	if (source.xml['a:pPr'] !== undefined) {
		field['a:pPr'] = clone(source.xml['a:pPr']);
	}
	field['a:t'] = segment.text;
	return { tag: 'a:fld', value: field };
}

/** Split segments into paragraphs at their paragraph breaks. */
function paragraphsOf(segments: readonly TextSegment[]): TextSegment[][] {
	const paragraphs: TextSegment[][] = [[]];
	for (const segment of segments) {
		if (segment.isParagraphBreak) {
			paragraphs.push([]);
		} else {
			paragraphs.at(-1)!.push(segment);
		}
	}
	return paragraphs;
}

/**
 * The `a:p` list for an edited cell, in CT_TextParagraph order
 * (`a:pPr?`, runs, `a:endParaRPr?`).
 */
export function buildRichCellParagraphs(
	segments: readonly TextSegment[],
	runPropertiesFromStyle: RunPropertiesFromStyle,
): XmlObject[] {
	let lastSource: XmlObject | undefined;
	return paragraphsOf(segments).map((paragraphSegments) => {
		const source = paragraphSegments.find((segment) => segment.tableParagraph)?.tableParagraph;
		const layout = source ?? lastSource;
		lastSource = layout;
		const paragraph: XmlObject = {};
		if (layout?.['a:pPr'] !== undefined) {
			paragraph['a:pPr'] = clone(layout['a:pPr']);
		}
		const children = paragraphSegments.flatMap((segment) => {
			const child = segmentChild(segment, runPropertiesFromStyle);
			return child ? [child] : [];
		});
		writeParagraphChildren(paragraph, children);
		if (layout?.['a:endParaRPr'] !== undefined) {
			paragraph['a:endParaRPr'] = clone(layout['a:endParaRPr']);
		}
		return paragraph;
	});
}
