import type { Block, DocumentModel, HeaderFooterContent } from '../model';
import { withComputedFieldRuns } from './computed-field-runs';
import { adaptDocumentModel } from './adapter';
import type { LayoutDocumentInput, LayoutTextFrame } from './input';
import type { TextMeasurer } from './measure';
import {
	headerFooterForPage,
	pageNumberValues,
	fieldDisplayText,
	mergeFieldPlaceholders,
	sectionPageCounts,
	type PageFieldContext,
} from './page-fields';
import { formatNoteNumber } from '../notes';
import { pageStyleRefs } from './page-style-refs';
import { layoutSections } from './page-flow';
import { pageFrames, placeInFrames } from './floats';
import { positionFloats } from './position-floats';
import { inFlow, type LayoutBlockBox, type LayoutFloatBox, type LayoutPageBox } from './result';
import type { LayoutReviewOptions } from './review-display';
import { createTextBoxLayouts, type TextBoxLayouts } from './text-box-layout';
import { refreshComputedFields } from '../computed-fields';
import { fieldName } from '../field-runs';
import { fieldResultSpans } from '../field-result-spans';
import { resolveRunFormatting } from '../run-formatting';

const DEFAULT_DISTANCE_PX = 48;

/** Uses the body layout for story paragraphs, tables and text boxes at the page's text width. */
export function flowStory(
	model: DocumentModel,
	blocks: Block[],
	widthPx: number,
	measurer: TextMeasurer,
	options: LayoutReviewOptions,
	note: (message: string) => void,
) {
	const {
		sections: _sections,
		footnoteSeparators: _separators,
		footnotes: _footnotes,
		endnotes: _endnotes,
		...base
	} = model;
	const doc = { ...base, blocks };
	const [story] = adaptDocumentModel(
		doc,
		note,
		options,
		storyTextBoxLayouts(doc, blocks, measurer, options, note),
	).sections;
	const input: LayoutDocumentInput = {
		sections: [
			{
				blocks: story?.blocks ?? [],
				page: {
					widthPx,
					heightPx: 1_000_000,
					marginTopPx: 0,
					marginRightPx: 0,
					marginBottomPx: 0,
					marginLeftPx: 0,
				},
				...(story?.addParagraphSpacing ? { addParagraphSpacing: true } : {}),
			},
		],
	};
	const result = layoutSections(withComputedFieldRuns(input, doc), measurer);
	for (const warning of result.approximations) note(warning);
	const storyPage = result.pages[0];
	const column = storyPage?.columns[0];
	const boxes = column?.blocks ?? [];
	const textFrames = new Map<string, LayoutTextFrame>();
	for (const block of story?.blocks ?? [])
		if (block.kind === 'paragraph' && block.textFrame) textFrames.set(block.id, block.textFrame);
	const joins = new Set(
		story?.blocks.flatMap((block) =>
			block.kind === 'paragraph' && block.borderJoinNext ? [block.id] : [],
		) ?? [],
	);
	for (const [index, box] of boxes.entries()) {
		if (!joins.has(box.blockId) || box.kind !== 'paragraph' || !box.frame) continue;
		const next = boxes[index + 1];
		if (
			next?.kind === 'paragraph' &&
			next.frame &&
			box.frame.leftPx === next.frame.leftPx &&
			box.frame.widthPx === next.frame.widthPx
		)
			box.frame = {
				...box.frame,
				heightPx: Math.max(box.heightPx, next.yPx - box.yPx) - (box.frame.topPx ?? 0),
			};
	}
	const heightPx = boxes.reduce(
		(bottom, box) => (inFlow(box) ? Math.max(bottom, box.yPx + box.heightPx) : bottom),
		0,
	);
	return { input, storyPage, column, heightPx, textFrames };
}

/** Move out-of-flow text frames from the measuring page to the physical page's anchors. */
function storyColumn(flow: ReturnType<typeof flowStory>, page: LayoutPageBox, storyTopPx: number) {
	const { column, storyPage, textFrames } = flow;
	if (!column || !storyPage || !textFrames.size) return column;
	const before = pageFrames(storyPage, column, { topPx: 0, heightPx: 0 });
	const after = pageFrames(page, column, { topPx: storyTopPx, heightPx: 0 });
	const blocks = column.blocks.map((box) => {
		if (box.kind !== 'paragraph' || !box.floating) return box;
		const anchor = textFrames.get(box.blockId);
		if (!anchor) return box;
		// The frame's own size cancels when taking the difference between aligned positions.
		const point = { ...anchor, widthPx: 0, heightPx: 0 };
		const local = placeInFrames(point, before);
		const physical = placeInFrames(point, after);
		return {
			...box,
			xPx: (box.xPx ?? 0) + physical.xPx - local.xPx - page.marginLeftPx,
			yPx: box.yPx + physical.yPx - local.yPx - storyTopPx,
		};
	});
	return { ...column, blocks };
}

export interface PlacedStory {
	kind: 'header' | 'footer';
	x: number;
	y: number;
	heightPx: number;
	widthPx: number;
	blocks: LayoutBlockBox[];
	source: Block[];
	content: HeaderFooterContent;
	floats: LayoutFloatBox[];
	distancePx: number;
}

function fillFields(blocks: Block[], values: PageFieldContext, model: DocumentModel): Block[] {
	const paragraph = (p: Extract<Block, { type: 'paragraph' }>) => {
		const paragraphValues: PageFieldContext = {
			...values,
			languageOf: (run) =>
				resolveRunFormatting(run, {
					runCatalog: model.characterStyles,
					paragraphCatalog: model.paragraphStyles,
					paragraphStyleId: p.style,
				}).language,
		};
		let changed = false;
		const source = model.mailMerge ? p.runs : mergeFieldPlaceholders(p.runs);
		if (source !== p.runs) changed = true;
		const runs = source.map((run) => {
			if (!run.field) return run;
			const text = fieldDisplayText(run, paragraphValues);
			// DATE/TIME are evaluated for the page's time, instead of again during adaptation.
			const name = fieldName(run.field.instr);
			if (name === 'DATE' || name === 'TIME') {
				changed = true;
				const { field: _field, ...plain } = run;
				return { ...plain, text };
			}
			if (text === run.text) return run;
			changed = true;
			return { ...run, text };
		});
		if (p.runs.some((run) => run.fieldChar === 'begin')) {
			for (const span of fieldResultSpans(p.runs, { resultless: true })) {
				if (
					span.locked ||
					span.removed ||
					(!span.resultless && (span.code === undefined || span.runs.length))
				)
					continue;
				const run = runs[span.to];
				if (!run) continue;
				const text = fieldDisplayText({ ...run, field: { instr: span.instr } }, paragraphValues);
				if (!text) continue;
				runs[span.to] = { ...run, text };
				changed = true;
			}
		}
		return changed ? { ...p, runs } : p;
	};
	let changed = false;
	const filled = blocks.map((block) => {
		if (block.type === 'paragraph') {
			const next = paragraph(block);
			changed ||= next !== block;
			return next;
		}
		let changedTable = false;
		const rows = block.rows.map((row) =>
			row.map((cell) => {
				const paragraphs = cell.paragraphs.map(paragraph);
				if (paragraphs.every((p, index) => p === cell.paragraphs[index])) return cell;
				changedTable = true;
				return { ...cell, paragraphs };
			}),
		);
		changed ||= changedTable;
		return changedTable ? { ...block, rows } : block;
	});
	return changed ? filled : blocks;
}

/** Resolves stories on demand. Repeat stories without changing field results reuse their layout. */
export function createPageStories(
	document: DocumentModel,
	pages: LayoutPageBox[],
	measurer: TextMeasurer,
	options: LayoutReviewOptions = {},
	note: (message: string) => void = () => {},
	now = new Date(),
) {
	const model = refreshComputedFields(document);
	const numbers = pageNumberValues(model, pages);
	const counts = sectionPageCounts(pages);
	const styleRefs = pageStyleRefs(model, pages);
	const cache = new WeakMap<Block[], Map<number, ReturnType<typeof flowStory>>>();
	return (
		index: number,
		page = pages[index],
		pageNumber = numbers[index] ?? index + 1,
	): PlacedStory[] => {
		if (!page) return [];
		const section = model.sections?.[page.sectionIndex];
		const widthPx = Math.max(1, page.widthPx - page.marginLeftPx - page.marginRightPx);
		const values = {
			page: formatNoteNumber(pageNumber, section?.pageNumbering?.format ?? 'decimal'),
			numPages: String(pages.length),
			sectionPages: String(counts.get(page.sectionIndex) ?? 1),
			now,
		};
		const placed: PlacedStory[] = [];
		for (const kind of ['header', 'footer'] as const) {
			const content = headerFooterForPage(model, page, pageNumber, `${kind}s`);
			if (!content) continue;
			const source = fillFields(styleRefs.content(content, index).blocks, values, model);
			let widths = cache.get(source);
			if (!widths) cache.set(source, (widths = new Map()));
			let flow = widths.get(widthPx);
			if (!flow) {
				flow = flowStory(model, source, widthPx, measurer, options, note);
				widths.set(widthPx, flow);
			}
			const distance =
				kind === 'header' ? section?.headerDistanceTwips : section?.footerDistanceTwips;
			const distancePx = distance === undefined ? DEFAULT_DISTANCE_PX : distance / 15;
			const y = kind === 'header' ? distancePx : page.heightPx - distancePx - flow.heightPx;
			const column = storyColumn(flow, page, y);
			const { floats: _bodyFloats, ...pageFrame } = page;
			const frame: LayoutPageBox = { ...pageFrame, columns: column ? [column] : [] };
			positionFloats(flow.input, [frame], y);
			placed.push({
				kind,
				x: page.marginLeftPx,
				y,
				widthPx,
				heightPx: flow.heightPx,
				blocks: column?.blocks ?? [],
				source,
				content,
				floats: frame.floats ?? [],
				distancePx,
			});
		}
		return placed;
	};
}

/**
 * The text boxes of one of `model`'s stories (its body, or a header or footer's blocks) laid out by
 * the page engine with the document's styles, linked chains included.
 */
export function storyTextBoxLayouts(
	model: DocumentModel,
	story: readonly Block[],
	measurer: TextMeasurer,
	options: LayoutReviewOptions = {},
	note: (message: string) => void = () => {},
): TextBoxLayouts {
	// The box's blocks adapt as a document of their own: the styles, theme and numbering are the
	// document's; its sections, notes and separators are not the box's.
	const {
		sections: _sections,
		footnoteSeparators: _separators,
		footnotes: _footnotes,
		endnotes: _endnotes,
		...base
	} = model;
	return createTextBoxLayouts(
		story,
		(blocks) => adaptDocumentModel({ ...base, blocks }, note, options).sections[0]?.blocks ?? [],
		measurer,
		note,
	);
}
