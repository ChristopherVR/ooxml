import { layoutDocument } from './layout';
import { SectionFlow } from './section-flow';
import type { PageCursor } from './page-cursor';
import type { LayoutBlock, LayoutDocumentInput } from './input';
import type { TextMeasurer } from './measure';
import type { LayoutBlockBox, LayoutPageBox, LayoutResult } from './result';

/**
 * How far past a paragraph a page break can depend on it: `keepNext` chains look ahead at most
 * this many blocks (`SectionFlow.requiredKeepHeight`), so pages that end earlier are unaffected.
 */
const LOOKAHEAD_BLOCKS = 66;

interface BlockState {
	/** The page the block's last piece sits on, and the fill position after it. */
	page: number;
	y: number;
}

interface Snapshot {
	/** Everything in the section but its blocks: page geometry and columns. */
	context: string;
	signatures: string[];
	state: BlockState[];
	/** Per page: the block its first box belongs to, and whether that box starts the block. */
	firstBlock: (number | undefined)[];
	startsBlock: boolean[];
	result: LayoutResult;
}

/** What the last pass did, for tests and diagnostics. */
export interface IncrementalStats {
	mode: 'full' | 'incremental' | 'unchanged';
	/** Blocks that went through pagination in that pass. */
	placed: number;
	/** The first page that was laid out again. */
	fromPage: number;
}

export interface IncrementalMemory {
	last?: Snapshot | undefined;
	stats?: IncrementalStats;
}

function paragraphBoxes(page: LayoutPageBox): LayoutBlockBox[] {
	return page.columns[0]?.blocks ?? [];
}

/**
 * A section of plain paragraphs in one column, with no floats, footnotes or tables. Page breaks
 * there depend only on the fill position, so a pass can resume at a page start and stop where the
 * layout matches the previous one again. Everything else takes the full pass.
 */
function eligible(input: LayoutDocumentInput): boolean {
	const [section] = input.sections;
	if (!section || input.sections.length !== 1) return false;
	if (
		(section.columns?.count ?? 1) > 1 ||
		section.footnoteColumns ||
		section.modernFootnoteSeparator ||
		section.legacyFootnoteColumns ||
		section.footnoteSeparators ||
		(section.verticalAlign && section.verticalAlign !== 'top')
	)
		return false;
	const ids = new Set<string>();
	return section.blocks.every((block: LayoutBlock) => {
		if (block.kind !== 'paragraph' || ids.has(block.id)) return false;
		ids.add(block.id);
		return !block.floats?.length && !block.footnotes?.length && !block.afterTableSectionBreak;
	});
}

function describePages(
	pages: readonly LayoutPageBox[],
	indexOf: ReadonlyMap<string, number>,
): Pick<Snapshot, 'firstBlock' | 'startsBlock'> {
	const firstBlock: (number | undefined)[] = [];
	const startsBlock: boolean[] = [];
	for (const page of pages) {
		const first = paragraphBoxes(page)[0];
		firstBlock.push(first ? indexOf.get(first.blockId) : undefined);
		const line = first?.kind === 'paragraph' ? first.lines[0] : undefined;
		startsBlock.push(line !== undefined && line.sourceStart === 0);
	}
	return { firstBlock, startsBlock };
}

function snapshot(
	blocks: readonly LayoutBlock[],
	context: string,
	signatures: string[],
	state: BlockState[],
	result: LayoutResult,
): Snapshot {
	const indexOf = new Map(blocks.map((block, index) => [block.id, index]));
	return { context, signatures, state, result, ...describePages(result.pages, indexOf) };
}

/** Lays out from scratch, recording where each block ended so a later pass can resume. */
function fullPass(
	input: LayoutDocumentInput,
	measurer: TextMeasurer,
	memory: IncrementalMemory,
	context: string,
	signatures: string[],
): LayoutResult {
	const [section] = input.sections;
	const blocks = section!.blocks;
	const notes = new Set<string>();
	const pages: LayoutPageBox[] = [];
	const state: BlockState[] = [];
	new SectionFlow(
		section!,
		pages,
		measurer,
		(message) => notes.add(message),
		0,
		undefined,
		undefined,
		undefined,
		{
			startIndex: 0,
			pageInSection: 0,
			afterBlock(index, cursor) {
				state[index] = { page: cursor.page.index, y: cursor.y };
				return false;
			},
		},
	).run();
	const result: LayoutResult = { pages, approximations: [...notes] };
	// Notes come from blocks a later pass might skip, so only note-free layouts are resumable.
	memory.last = notes.size ? undefined : snapshot(blocks, context, signatures, state, result);
	memory.stats = { mode: 'full', placed: blocks.length, fromPage: 0 };
	return result;
}

function sameBoxes(a: readonly LayoutBlockBox[], b: readonly LayoutBlockBox[]): boolean {
	return (
		a.length <= b.length &&
		a.every((box, index) => JSON.stringify(box) === JSON.stringify(b[index]))
	);
}

/**
 * Paginates like `layoutDocument`, but reuses the previous pass when the input is a plain
 * single-section document: pagination resumes at a page start before the first changed paragraph
 * and stops as soon as a page matches the previous layout again, taking the old pages from there.
 * Results are identical to a full pass; anything outside that shape runs the full pass.
 */
export function layoutDocumentIncremental(
	input: LayoutDocumentInput,
	measurer: TextMeasurer,
	memory: IncrementalMemory,
): LayoutResult {
	if (!eligible(input)) {
		memory.last = undefined;
		memory.stats = { mode: 'full', placed: 0, fromPage: 0 };
		return layoutDocument(input, measurer);
	}
	const section = input.sections[0]!;
	const blocks = section.blocks;
	const signatures = blocks.map((block) => JSON.stringify(block));
	const context = JSON.stringify({ ...section, blocks: undefined });
	const previous = memory.last;
	if (!previous || previous.context !== context)
		return fullPass(input, measurer, memory, context, signatures);

	const oldCount = previous.signatures.length;
	let prefix = 0;
	while (
		prefix < blocks.length &&
		prefix < oldCount &&
		signatures[prefix] === previous.signatures[prefix]
	)
		prefix++;
	if (prefix === blocks.length && prefix === oldCount) {
		memory.stats = { mode: 'unchanged', placed: 0, fromPage: previous.result.pages.length };
		return previous.result;
	}
	let suffix = 0;
	while (
		suffix < blocks.length - prefix &&
		suffix < oldCount - prefix &&
		signatures[blocks.length - 1 - suffix] === previous.signatures[oldCount - 1 - suffix]
	)
		suffix++;
	const delta = oldCount - blocks.length;

	// The last page start that no edit can reach back to, with its first block placed whole.
	let page = 0;
	for (let index = previous.firstBlock.length - 1; index > 0; index--) {
		const first = previous.firstBlock[index];
		if (first !== undefined && previous.startsBlock[index] && first <= prefix - LOOKAHEAD_BLOCKS) {
			page = index;
			break;
		}
	}
	const start = page === 0 ? 0 : previous.firstBlock[page]!;
	const old = previous.result.pages;
	const pages = old.slice(0, page);
	const state = previous.state.slice(0, start);
	const notes = new Set<string>();
	let placed = 0;
	let stop: { oldPage: number; at: number } | undefined;
	new SectionFlow(
		section,
		pages,
		measurer,
		(message) => notes.add(message),
		0,
		undefined,
		undefined,
		undefined,
		{
			startIndex: start,
			pageInSection: page,
			afterBlock(index, cursor: PageCursor) {
				placed++;
				state[index] = { page: cursor.page.index, y: cursor.y };
				const oldIndex = index + delta;
				if (index < blocks.length - suffix || oldIndex < 0 || oldIndex >= oldCount) return false;
				const before = previous.state[oldIndex];
				if (before?.y !== cursor.y) return false;
				const oldPage = old[before.page];
				if (!oldPage || !sameBoxes(paragraphBoxes(cursor.page), paragraphBoxes(oldPage)))
					return false;
				stop = { oldPage: before.page, at: index };
				return true;
			},
		},
	).run();
	if (notes.size) return fullPass(input, measurer, memory, context, signatures);

	if (stop) {
		// The current page continues exactly as it did before: take it and every later page.
		const shift = pages.length - 1 - stop.oldPage;
		pages.pop();
		for (const kept of old.slice(stop.oldPage))
			pages.push(
				shift === 0
					? kept
					: { ...kept, index: kept.index + shift, pageInSection: kept.pageInSection + shift },
			);
		for (let index = stop.at + 1; index < blocks.length; index++) {
			const before = previous.state[index + delta]!;
			state[index] = { page: before.page + shift, y: before.y };
		}
	}
	const result: LayoutResult = { pages, approximations: [] };
	memory.last = snapshot(blocks, context, signatures, state, result);
	memory.stats = { mode: 'incremental', placed, fromPage: page };
	return result;
}
