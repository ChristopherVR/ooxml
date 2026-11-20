import type { Block, InlineImage, TextBoxContent } from '../index';
import { EMU_PER_PIXEL } from '../../units/constants';
import { layoutSections } from './page-flow';
import type { LayoutBlock, LayoutColumns, LayoutRun } from './input';
import { inFlow, type LayoutBlockBox } from './result';
import { evaluatePresetShape, lookupPresetShape } from '../../geometry/preset-shape-evaluator';
import { shapeAppearance } from '../shape-svg';
import type { TextMeasurer } from './measure';

// Text boxes laid out by the page engine: a box's paragraphs are measured and broken into lines by
// the same paragraph flow as the body (`layoutSections`), inside the shape's text rectangle (the
// preset's text rectangle less the `wps:bodyPr` insets and half the outline). A linked chain is one section whose columns are its boxes,
// so a paragraph that overflows one box reflows at the next box's width. Word 16 shows the overflow
// of the last box clipped at its text rectangle, so the last box takes every remaining line.

/** A text box's paragraphs laid out inside its shape. */
export interface TextBoxLayout {
	/** The shape's drawn size: its extent, or the size `spAutoFit` or `wrap="none"` gives it. */
	widthPx: number;
	heightPx: number;
	/**
	 * The text rectangle, in the shape's own coordinates: the preset's text rectangle less the
	 * insets and half the outline.
	 */
	frame: { xPx: number; yPx: number; widthPx: number; heightPx: number };
	/**
	 * `vert` (and the East Asian and WordArt vertical kinds) turns the text a quarter turn clockwise,
	 * so lines run down and the first is at the right; `vert270` a quarter turn anticlockwise.
	 */
	rotation?: 90 | 270;
	/**
	 * The laid-out blocks in text coordinates: x runs along a line (the frame's width, or its height
	 * when rotated), y across lines, already moved for the box's top, middle or bottom anchor.
	 */
	blocks: LayoutBlockBox[];
}

/** Lays out the text boxes of one story (the body, a header or a footer), linked chains included. */
export interface TextBoxLayouts {
	layout(
		image: Pick<InlineImage, 'shape' | 'textBoxContent' | 'widthPx' | 'heightPx'>,
	): TextBoxLayout | undefined;
}

/** Adapts a text box's blocks to engine input with the document's styles. */
export type AdaptTextBoxBlocks = (blocks: Block[]) => LayoutBlock[];

/** `wps:bodyPr` insets default to 0.1 inch left and right, 0.05 inch top and bottom. */
const DEFAULT_INSETS_EMU = { left: 91440, top: 45720, right: 91440, bottom: 45720 };
/** Height of the column that takes every remaining line (the last box of a chain). */
const UNBOUNDED_PX = 1_000_000;
export const VERTICAL_TEXT_NOTE =
	'East Asian and WordArt vertical text in text boxes is drawn turned a quarter clockwise, without upright East Asian characters or stacked letters.';

type Sized = Pick<InlineImage, 'shape' | 'textBoxContent' | 'widthPx' | 'heightPx'>;

function rotationOf(vertical: string | undefined): TextBoxLayout['rotation'] {
	if (!vertical || vertical === 'horz') return undefined;
	return vertical === 'vert270' ? 270 : 90;
}

/**
 * Space between the shape's edges and its text: the preset's text rectangle (all of a `rect`), less
 * the `wps:bodyPr` insets and half the outline's width, which Word 16 keeps text clear of (measured:
 * a 6 pt outline wraps a line that fits inside a box without one).
 */
function paddingOf(box: Sized, content: TextBoxContent) {
	const px = (emu: number | undefined, fallback: number) => (emu ?? fallback) / EMU_PER_PIXEL;
	const spec = box.shape;
	const area =
		spec && !spec.customGeometryXml && lookupPresetShape(spec.preset)
			? evaluatePresetShape(spec.preset, box.widthPx, box.heightPx, spec.adjustments)?.textRect
			: undefined;
	const look = spec && shapeAppearance(spec);
	const half = look && look.stroke !== 'none' ? look.strokeWidthPx / 2 : 0;
	return {
		left: (area?.l ?? 0) + px(content.insetLeftEmu, DEFAULT_INSETS_EMU.left) + half,
		top: (area?.t ?? 0) + px(content.insetTopEmu, DEFAULT_INSETS_EMU.top) + half,
		right:
			(area ? box.widthPx - area.r : 0) +
			px(content.insetRightEmu, DEFAULT_INSETS_EMU.right) +
			half,
		bottom:
			(area ? box.heightPx - area.b : 0) +
			px(content.insetBottomEmu, DEFAULT_INSETS_EMU.bottom) +
			half,
	};
}

/** Word draws a page or column break inside a text box as a line break (measured, Word 16). */
function breaksAsLines(block: LayoutBlock): LayoutBlock {
	if (block.kind !== 'paragraph' || !block.runs.some((run) => run.breakAfter)) return block;
	return {
		...block,
		runs: block.runs.map((run): LayoutRun => {
			if (!run.breakAfter) return run;
			const { breakAfter: _break, ...rest } = run;
			return { ...rest, text: `${run.text}\n` };
		}),
	};
}

/** Lays `blocks` out in one section of columns; returns each column's placed blocks. */
function flow(
	blocks: LayoutBlock[],
	columns: { widthPx: number; heightPx?: number }[],
	measurer: TextMeasurer,
): LayoutBlockBox[][] {
	const layoutColumns: LayoutColumns = {
		count: columns.length,
		gapPx: 0,
		widths: columns.map((column) => ({ ...column, gapPx: 0 })),
	};
	const { pages } = layoutSections(
		{
			sections: [
				{
					page: {
						widthPx: columns.reduce((sum, column) => sum + column.widthPx, 0),
						heightPx: UNBOUNDED_PX,
						marginTopPx: 0,
						marginRightPx: 0,
						marginBottomPx: 0,
						marginLeftPx: 0,
					},
					columns: layoutColumns,
					blocks,
				},
			],
		},
		measurer,
	);
	return columns.map((_, index) => pages[0]?.columns[index]?.blocks ?? []);
}

/** The widest line of `blocks` laid out without wrapping: what a `wrap="none"` box is sized to. */
function naturalLength(blocks: LayoutBlock[], measurer: TextMeasurer): number {
	const leftAligned = blocks.map((block): LayoutBlock => {
		if (block.kind !== 'paragraph' || !block.align) return block;
		const { align: _align, ...rest } = block;
		return rest;
	});
	const [placed = []] = flow(leftAligned, [{ widthPx: UNBOUNDED_PX }], measurer);
	let widest = 1;
	for (const block of placed) {
		if (block.kind === 'table') widest = Math.max(widest, (block.xPx ?? 0) + (block.widthPx ?? 0));
		else
			for (const line of block.lines)
				for (const fragment of line.fragments)
					widest = Math.max(widest, fragment.xPx + fragment.widthPx);
	}
	return widest;
}

/** Lays out a chain of linked boxes (one box when unlinked), its text taken from `blocks`. */
export function layoutTextBoxChain(
	boxes: Sized[],
	blocks: LayoutBlock[],
	measurer: TextMeasurer,
): TextBoxLayout[] {
	const flowing = blocks.map(breaksAsLines);
	const single = boxes.length === 1;
	const geometry = boxes.map((box, index) => {
		const content = box.textBoxContent ?? { blocks: [] };
		const inset = paddingOf(box, content);
		const rotation = rotationOf(content.vertical);
		const frameWidth = Math.max(1, box.widthPx - inset.left - inset.right);
		const frameHeight = Math.max(1, box.heightPx - inset.top - inset.bottom);
		// A box that does not wrap is as wide as its longest line (Word resizes it on display).
		const lengthPx =
			single && content.wrap === 'none'
				? naturalLength(flowing, measurer)
				: rotation
					? frameHeight
					: frameWidth;
		return {
			box,
			content,
			inset,
			rotation,
			lengthPx,
			depthPx: rotation ? frameWidth : frameHeight,
			last: index === boxes.length - 1,
		};
	});
	const placed = flow(
		flowing,
		geometry.map((item) => ({
			widthPx: item.lengthPx,
			...(item.last ? {} : { heightPx: item.depthPx }),
		})),
		measurer,
	);
	return geometry.map((item, index): TextBoxLayout => {
		const blocksIn = placed[index] ?? [];
		const used = Math.max(0, ...blocksIn.filter(inFlow).map((block) => block.yPx + block.heightPx));
		// `spAutoFit`: the shape's height (its width for vertical text) follows its text, as Word draws it.
		const depthPx = single && item.content.autofit === 'shape' ? used : item.depthPx;
		const free = depthPx - used;
		const shift =
			free <= 0
				? 0
				: item.content.anchor === 'ctr'
					? free / 2
					: item.content.anchor === 'b'
						? free
						: 0;
		const { inset, rotation, lengthPx } = item;
		const frame = {
			xPx: inset.left,
			yPx: inset.top,
			widthPx: rotation ? depthPx : lengthPx,
			heightPx: rotation ? lengthPx : depthPx,
		};
		// Only autofit and `wrap="none"` resize the shape; otherwise it keeps its extent.
		const resized = depthPx !== item.depthPx || item.content.wrap === 'none';
		return {
			widthPx: resized ? frame.widthPx + inset.left + inset.right : item.box.widthPx,
			heightPx: resized ? frame.heightPx + inset.top + inset.bottom : item.box.heightPx,
			frame,
			...(rotation ? { rotation } : {}),
			blocks: shift ? blocksIn.map((block) => ({ ...block, yPx: block.yPx + shift })) : blocksIn,
		};
	});
}

function storyTextBoxes(blocks: readonly Block[], found: Sized[]): Sized[] {
	for (const block of blocks) {
		if (block.type === 'table') {
			for (const row of block.rows)
				for (const cell of row) storyTextBoxes(cell.sourceBlocks ?? cell.paragraphs, found);
			continue;
		}
		for (const run of block.runs) if (run.image?.textBoxContent) found.push(run.image);
	}
	return found;
}

/**
 * The text box layouts of one story: each box asked for is laid out once, a linked chain (found in
 * `story` by its `wps:txbx` id) together, its text flowing from box to box in `seq` order.
 */
export function createTextBoxLayouts(
	story: readonly Block[],
	adapt: AdaptTextBoxBlocks,
	measurer: TextMeasurer,
	note: (message: string) => void = () => {},
): TextBoxLayouts {
	const chains = new Map<number, Sized[]>();
	for (const image of storyTextBoxes(story, [])) {
		const link = image.textBoxContent?.link;
		if (!link) continue;
		const chain = chains.get(link.id) ?? [];
		if (!chain.some((box) => box.textBoxContent?.link?.seq === link.seq)) chain.push(image);
		chains.set(link.id, chain);
	}
	for (const chain of chains.values())
		chain.sort((a, b) => (a.textBoxContent?.link?.seq ?? 0) - (b.textBoxContent?.link?.seq ?? 0));
	const singles = new WeakMap<TextBoxContent, TextBoxLayout>();
	const linked = new Map<number, TextBoxLayout[]>();
	const laidOut = (boxes: Sized[]) => {
		const source = boxes.find((box) => box.textBoxContent?.blocks.length)?.textBoxContent;
		if (
			boxes.some(
				(box) =>
					rotationOf(box.textBoxContent?.vertical) &&
					!/^vert(270)?$/u.test(box.textBoxContent?.vertical ?? ''),
			)
		)
			note(VERTICAL_TEXT_NOTE);
		return layoutTextBoxChain(boxes, source ? adapt(source.blocks) : [], measurer);
	};
	return {
		layout(image) {
			const content = image.textBoxContent;
			if (!content) return undefined;
			const chain = content.link && chains.get(content.link.id);
			if (chain && chain.length > 1 && content.link) {
				let layouts = linked.get(content.link.id);
				if (!layouts) linked.set(content.link.id, (layouts = laidOut(chain)));
				const seq = content.link.seq;
				return layouts[chain.findIndex((box) => box.textBoxContent?.link?.seq === seq)];
			}
			let layout = singles.get(content);
			if (!layout) singles.set(content, (layout = laidOut([image])[0]!));
			return layout;
		},
	};
}
