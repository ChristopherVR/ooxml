/**
 * Puts each page's paragraphs, tables and pictures in reading order (`flow.ts`), joins a paragraph
 * or a table that a page break cut, and lays the pages out as sections: one per run of pages with
 * the same size, and a continuous one wherever the column layout changes.
 */
import { pageFlow, itemBottom, itemTop, type ColumnLayout, type FlowItem } from './flow';
import type { Line } from './lines';
import type { PictureBlock, SectionLayout } from './package';
import { continuesOnNextPage, type TextArea } from './paragraphs';
import type { RepeatedStory } from './stories';
import { joinTables, type TableBlock } from './tables';

export interface AssemblePage {
	width: number;
	height: number;
	lines: Line[];
	pictures: (PictureBlock & { x: number })[];
	tables?: TableBlock[];
}

function sectionFor(
	pages: AssemblePage[],
	area: TextArea,
	repeated: RepeatedStory[],
	start: number,
): SectionLayout {
	const page = pages[0]!;
	// Text lines and table frames (whose text left the lines) bound the text area.
	const boxes = pages.flatMap((item) => [...item.lines, ...(item.tables ?? [])]);
	const clamp = (value: number, max: number) => Math.min(max, Math.max(18, value));
	const header = repeated.find((story) => story.kind === 'header');
	const footer = repeated.find((story) => story.kind === 'footer');
	return {
		width: page.width,
		height: page.height,
		margins: {
			top: clamp(boxes.length ? Math.min(...boxes.map((box) => box.top)) - 2 : 72, page.height / 3),
			bottom: clamp(
				boxes.length ? page.height - Math.max(...boxes.map((box) => box.bottom)) - 2 : 72,
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

const layoutKey = (layout: ColumnLayout | undefined) =>
	layout ? `${layout.count}/${Math.round(layout.space)}/${layout.widths?.join(',') ?? ''}` : '1';

function withLayout(section: SectionLayout, layout: ColumnLayout | undefined): SectionLayout {
	const { columns: _columns, ...rest } = section;
	return layout ? { ...rest, columns: layout } : rest;
}

export function assemble(
	pages: AssemblePage[],
	area: TextArea,
	bodySize: number,
	repeated: RepeatedStory[],
): { blocks: FlowItem[]; sections: SectionLayout[] } {
	const blocks: FlowItem[] = [];
	const sections: SectionLayout[] = [];
	const textWidth = area.right - area.left;
	const sizeOf = (page: AssemblePage) => `${Math.round(page.width)}x${Math.round(page.height)}`;
	let currentLayout: ColumnLayout | undefined;
	pages.forEach((page, index) => {
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
		const { items, layouts } = pageFlow(page, area, bodySize);
		// Space around pictures and tables comes from where they sit, since paragraphs measure only text.
		items.forEach((item, at) => {
			const previous = items[at - 1];
			if (!previous || (item.kind === 'paragraph' && previous.kind === 'paragraph')) return;
			if (item.kind === 'table') return;
			const gap = Math.max(0, itemTop(item) - itemBottom(previous) - 2);
			item.spaceBefore = gap < 2 ? 0 : gap;
		});
		// A paragraph or a table cut by the page break goes on as one.
		const last = blocks[blocks.length - 1];
		const first = items[0];
		const sameLayout = layoutKey(currentLayout) === layoutKey(layouts[0]);
		if (last && first && sameLayout) {
			let joined = false;
			if (
				last.kind === 'paragraph' &&
				first.kind === 'paragraph' &&
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
				joined = true;
			} else if (last.kind === 'table' && first.kind === 'table') joined = joinTables(last, first);
			if (joined) {
				items.shift();
				layouts.shift();
			}
		}
		if (index === 0 || sizeOf(pages[index - 1]!) !== sizeOf(page)) {
			const run = [page];
			for (
				let next = index + 1;
				next < pages.length && sizeOf(pages[next]!) === sizeOf(page);
				next++
			)
				run.push(pages[next]!);
			currentLayout = layouts[0];
			sections.push(withLayout(sectionFor(run, area, repeated, blocks.length), currentLayout));
		}
		items.forEach((item, at) => {
			const layout = layouts[at];
			if (layoutKey(layout) !== layoutKey(currentLayout)) {
				const { continuous: _continuous, ...base } = sections[sections.length - 1]!;
				sections.push({ ...withLayout(base, layout), start: blocks.length, continuous: true });
				currentLayout = layout;
			}
			blocks.push(item);
		});
	});
	return { blocks, sections };
}
