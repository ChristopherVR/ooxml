/**
 * Puts each page's paragraphs and pictures in reading order, joins a paragraph that a page break
 * cut, and lays the pages out as sections (one per run of pages with the same size), with the
 * margins the text occupies.
 */
import type { Line } from './lines';
import type { PictureBlock, SectionLayout } from './package';
import {
	buildParagraphs,
	continuesOnNextPage,
	type ParagraphBlock,
	type TextArea,
} from './paragraphs';
import type { RepeatedStory } from './stories';

export interface AssemblePage {
	width: number;
	height: number;
	lines: Line[];
	pictures: (PictureBlock & { x: number })[];
}

const top = (block: ParagraphBlock | PictureBlock) =>
	block.kind === 'picture' ? block.top : block.lines[0]!.top;
const bottom = (block: ParagraphBlock | PictureBlock) => {
	if (block.kind === 'picture') return block.top + block.height;
	const last = block.lines[block.lines.length - 1]!;
	return last.bottom;
};

function sectionFor(
	pages: AssemblePage[],
	area: TextArea,
	repeated: RepeatedStory[],
	start: number,
): SectionLayout {
	const page = pages[0]!;
	const lines = pages.flatMap((item) => item.lines);
	const clamp = (value: number, max: number) => Math.min(max, Math.max(18, value));
	const header = repeated.find((story) => story.kind === 'header');
	const footer = repeated.find((story) => story.kind === 'footer');
	return {
		width: page.width,
		height: page.height,
		margins: {
			top: clamp(
				lines.length ? Math.min(...lines.map((line) => line.top)) - 2 : 72,
				page.height / 3,
			),
			bottom: clamp(
				lines.length ? page.height - Math.max(...lines.map((line) => line.bottom)) - 2 : 72,
				page.height / 3,
			),
			left: clamp(area.left, page.width / 3),
			right: clamp(page.width - area.right, page.width / 3),
		},
		header: header ? Math.max(12, header.distance) : 36,
		footer: footer ? Math.max(12, footer.distance) : 36,
		start,
	};
}

export function assemble(
	pages: AssemblePage[],
	area: TextArea,
	bodySize: number,
	repeated: RepeatedStory[],
): { blocks: (ParagraphBlock | PictureBlock)[]; sections: SectionLayout[] } {
	const blocks: (ParagraphBlock | PictureBlock)[] = [];
	const sections: SectionLayout[] = [];
	const textWidth = area.right - area.left;
	const sizeOf = (page: AssemblePage) => `${Math.round(page.width)}x${Math.round(page.height)}`;
	pages.forEach((page, index) => {
		if (index === 0 || sizeOf(pages[index - 1]!) !== sizeOf(page)) {
			const run = [page];
			for (
				let next = index + 1;
				next < pages.length && sizeOf(pages[next]!) === sizeOf(page);
				next++
			)
				run.push(pages[next]!);
			sections.push(sectionFor(run, area, repeated, blocks.length));
		}
		for (const picture of page.pictures) {
			if (picture.width > textWidth) {
				picture.height *= textWidth / picture.width;
				picture.width = textWidth;
			}
			const before = picture.x - area.left;
			const after = area.right - (picture.x + picture.width);
			picture.align =
				before > 6 && Math.abs(before - after) < 8
					? 'center'
					: before > 6 && after < 6
						? 'right'
						: 'left';
		}
		const items = [...buildParagraphs(page.lines, area, bodySize), ...page.pictures].sort(
			(a, b) => top(a) - top(b),
		);
		// Space around pictures comes from where they sit, since paragraphs measure only text.
		items.forEach((item, at) => {
			const previous = items[at - 1];
			if (!previous || (item.kind === 'paragraph' && previous.kind === 'paragraph')) return;
			const gap = Math.max(0, top(item) - bottom(previous) - 2);
			if (item.kind === 'picture') item.spaceBefore = gap < 2 ? 0 : gap;
			else item.spaceBefore = gap < 2 ? 0 : gap;
		});
		// A paragraph cut by the page break goes on as one paragraph.
		const last = blocks[blocks.length - 1];
		const first = items[0];
		if (
			last?.kind === 'paragraph' &&
			first?.kind === 'paragraph' &&
			continuesOnNextPage(last, first)
		) {
			const tail = last.fragments[last.fragments.length - 1];
			if (tail) {
				const hyphen =
					/[-\u00ad]$/.test(tail.text) && /^\s*[a-z\u00df-\u00ff]/.test(first.lines[0]!.text);
				tail.text = hyphen ? tail.text.replace(/[-\u00ad]$/, '') : `${tail.text} `;
			}
			last.fragments.push(...first.fragments);
			last.lines.push(...first.lines);
			items.shift();
		}
		blocks.push(...items);
	});
	return { blocks, sections };
}
