import type { LayoutFootnoteBox, LayoutParagraphBox } from './result';
/** Split a note at a complete line boundary, retaining its identity and source ranges. */
export function splitNote(
	note: LayoutFootnoteBox,
	budget: number,
): {
	before?: LayoutFootnoteBox;
	after?: LayoutFootnoteBox;
} {
	if (note.heightPx <= budget + 1e-6) return { before: note };
	const before: LayoutParagraphBox[] = [];
	const after: LayoutParagraphBox[] = [];
	let origin: number | undefined;
	const sourceTokenStarts = { ...note.sourceTokenStarts };
	for (const paragraph of note.paragraphs) {
		if (origin !== undefined) {
			after.push({ ...paragraph, yPx: paragraph.yPx - origin });
			continue;
		}
		if (paragraph.yPx + paragraph.heightPx <= budget + 1e-6) {
			before.push(paragraph);
			continue;
		}
		const count = paragraph.lines.filter(
			(line) => paragraph.yPx + line.yPx + line.heightPx <= budget + 1e-6,
		).length;
		const head = paragraph.lines.slice(0, count);
		const tail = paragraph.lines.slice(count);
		if (head.length) {
			const bottom = head.at(-1)!.yPx + head.at(-1)!.heightPx;
			before.push({ ...paragraph, heightPx: bottom, lines: head });
		}
		const first = tail[0];
		if (head.length)
			sourceTokenStarts[paragraph.blockId] =
				head.at(-1)!.nextToken ?? sourceTokenStarts[paragraph.blockId] ?? 0;
		origin = paragraph.yPx + (first?.yPx ?? 0);
		after.push({
			...paragraph,
			yPx: 0,
			heightPx: Math.max(0, paragraph.heightPx - (first?.yPx ?? 0)),
			lines: tail.map((line) => ({ ...line, yPx: line.yPx - (first?.yPx ?? 0) })),
		});
	}
	if (!before.length) return { after: note };
	const last = before.at(-1)!;
	return {
		before: { ...note, heightPx: last.yPx + last.heightPx, paragraphs: before },
		...(after.length
			? {
					after: {
						...note,
						yPx: 0,
						heightPx: Math.max(0, note.heightPx - (origin ?? 0)),
						paragraphs: after,
						continued: true,
						sourceTokenStarts,
					},
				}
			: {}),
	};
}
