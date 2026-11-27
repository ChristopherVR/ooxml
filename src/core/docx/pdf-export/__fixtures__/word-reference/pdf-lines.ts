/**
 * The text lines of each page of a PDF, top to bottom (the reader's y grows down the page), for comparing line and page breaks with
 * Word's own Save as PDF: spans are grouped by baseline (a superscript joins its line), ordered
 * by x, and each line is its text with the white space removed, so the comparison does not depend
 * on how a writer splits words into spans or writes the spaces between them.
 */
import { openPdf } from '../../../../pdf/index';

export async function pdfLines(bytes: Uint8Array): Promise<string[][]> {
	const reader = await openPdf(bytes);
	const pages: string[][] = [];
	for (let index = 0; index < reader.pageCount; index++) {
		const { spans } = await reader.page(index);
		const visible = spans
			.filter((span) => !span.invisible && span.text.trim())
			.sort((a, b) => a.y - b.y || a.x - b.x);
		const lines: { y: number; size: number; spans: typeof visible }[] = [];
		for (const span of visible) {
			const line = lines.find(
				(candidate) =>
					Math.abs(candidate.y - span.y) < 0.45 * Math.max(candidate.size, span.fontSize),
			);
			if (line) {
				line.spans.push(span);
				if (span.fontSize > line.size) {
					line.size = span.fontSize;
					line.y = span.y;
				}
			} else lines.push({ y: span.y, size: span.fontSize, spans: [span] });
		}
		pages.push(
			lines
				.sort((a, b) => a.y - b.y)
				.map((line) =>
					line.spans
						.sort((a, b) => a.x - b.x)
						.map((span) => span.text)
						.join('')
						.replace(/\s+/g, ''),
				),
		);
	}
	return pages;
}
