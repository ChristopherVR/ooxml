import type { VisioParagraph, VisioText, VisioTextRun } from 'ooxml-core/visio';
import {
	wrapText,
	measureBrowserText,
	createTextLayoutBudget,
	TextLayoutBudgetError,
	type TextLayoutBudget,
	type MeasureText,
	type TextLine,
} from './text-layout.js';
export interface PositionedLine extends TextLine {
	x: number;
	y: number;
	availableWidth: number;
	align: VisioParagraph['horizontalAlign'];
	direction: 'ltr' | 'rtl';
	last: boolean;
	bullet?: NonNullable<VisioParagraph['bullet']> & { x: number };
}
export interface ParagraphLayout {
	lines: PositionedLine[];
	height: number;
}
export function layoutParagraphs(
	text: VisioText,
	measure: MeasureText = measureBrowserText,
	budget: TextLayoutBudget = createTextLayoutBudget(),
): ParagraphLayout {
	const scope = budget.forShape();
	const runs = text.runs.length
		? text.runs
		: [
				{
					text: text.plainText,
					fontFamily: text.fontFamily,
					fontSize: text.fontSize,
					color: text.color,
					bold: false,
					italic: false,
					underline: false,
				},
			];
	if (runs.length > 10_000 || text.plainText.length > scope.limits.maxCharacters)
		throw new TextLayoutBudgetError('Text layout exceeds its per-shape run/character limit.');
	// One prefix index replaces an O(paragraphs * runs) scan from the beginning of each paragraph.
	const offsets = [0];
	for (const run of runs) {
		const next = offsets[offsets.length - 1]! + run.text.length;
		if (next > scope.limits.maxCharacters)
			throw new TextLayoutBudgetError('Text layout exceeds its per-shape character limit.');
		offsets.push(next);
	}
	const paragraphs: VisioParagraph[] = text.paragraphs?.length
		? text.paragraphs
		: [
				{
					start: 0,
					end: text.plainText.length,
					horizontalAlign: text.horizontalAlign,
					indentLeft: 0,
					indentRight: 0,
					indentFirst: 0,
					spaceBefore: 0,
					spaceAfter: 0,
					lineSpacing: { kind: 'multiple', value: 1.2 },
					direction: 'ltr',
				},
			];
	if (paragraphs.length > 5000)
		throw new TextLayoutBudgetError('Text layout exceeds the safe paragraph limit.');
	let previousEnd = 0,
		coveredCharacters = 0;
	for (const paragraph of paragraphs) {
		if (
			!Number.isSafeInteger(paragraph.start) ||
			!Number.isSafeInteger(paragraph.end) ||
			paragraph.start < previousEnd ||
			paragraph.end < paragraph.start ||
			paragraph.end > text.plainText.length
		)
			throw new TextLayoutBudgetError(
				'Paragraph ranges must be ordered, non-overlapping and within the text.',
			);
		previousEnd = paragraph.end;
		coveredCharacters += paragraph.end - paragraph.start;
	}
	// Paragraph separators and uncovered ranges still consume budget even though runs omit them.
	scope.reserveCharacters(text.plainText.length - coveredCharacters);
	const lines: PositionedLine[] = [];
	let y = 0;
	for (const paragraph of paragraphs) {
		const source = sliceRuns(runs, offsets, paragraph.start, paragraph.end);
		const left = text.margins.left + paragraph.indentLeft,
			right = text.width - text.margins.right - paragraph.indentRight;
		const wrapped = wrapText(
			source,
			Math.max(0.001, right - left),
			measure,
			paragraph.indentFirst,
			scope,
		);
		y += paragraph.spaceBefore;
		for (const [index, line] of wrapped.entries()) {
			if (lines.length >= 5000)
				throw new TextLayoutBudgetError('Text layout exceeds the safe line limit.');
			const indent = index === 0 ? paragraph.indentFirst : 0;
			const minX = left + indent;
			const x =
				paragraph.horizontalAlign === 'center'
					? (minX + right) / 2
					: paragraph.horizontalAlign === 'right'
						? right
						: minX;
			const largestFont =
				line.runs.reduce((size, run) => Math.max(size, run.fontSize), 0) || text.fontSize;
			const height = Math.max(
				0.001,
				paragraph.lineSpacing.kind === 'exact'
					? paragraph.lineSpacing.value
					: largestFont * paragraph.lineSpacing.value,
			);
			lines.push({
				...line,
				height,
				x,
				y: y + largestFont * 0.8,
				availableWidth: Math.max(0.001, right - minX),
				align: paragraph.horizontalAlign,
				direction: paragraph.direction,
				last: index === wrapped.length - 1,
				...(index === 0 && paragraph.bullet
					? { bullet: { ...paragraph.bullet, x: left - paragraph.bullet.offset } }
					: {}),
			});
			y += height;
		}
		y += paragraph.spaceAfter;
	}
	return { lines, height: y };
}
function sliceRuns(
	runs: VisioTextRun[],
	offsets: number[],
	start: number,
	end: number,
): VisioTextRun[] {
	if (start === end) return [];
	let lo = 0,
		hi = runs.length;
	while (lo < hi) {
		const mid = (lo + hi) >>> 1;
		if (offsets[mid + 1]! <= start) lo = mid + 1;
		else hi = mid;
	}
	const result: VisioTextRun[] = [];
	for (let index = lo; index < runs.length && offsets[index]! < end; index++) {
		const run = runs[index]!,
			offset = offsets[index]!;
		const text = run.text.slice(Math.max(0, start - offset), Math.max(0, end - offset));
		if (text) result.push({ ...run, text });
	}
	return result;
}
