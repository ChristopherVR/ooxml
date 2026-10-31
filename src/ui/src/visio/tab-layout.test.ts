import type { VisioText } from 'ooxml-core/visio';
import { describe, expect, it } from 'vitest';
import { layoutParagraphs } from './paragraph-layout';
import { renderText } from './render-text';
import { layoutTabbedLine } from './tab-layout';

const measure = (text: string) => text.length * 0.1;
const run = (text: string) => ({
	text,
	fontFamily: 'Calibri',
	fontSize: 0.14,
	color: '#000000',
	bold: false,
	italic: false,
	underline: false,
});
const text = (
	plain: string,
	tabs: NonNullable<VisioText['paragraphs']>[number]['tabs'],
	align: 'left' | 'center' | 'right' | 'justify' = 'left',
): VisioText =>
	({
		plainText: plain,
		runs: [run(plain)],
		paragraphs: [
			{
				start: 0,
				end: plain.length,
				horizontalAlign: align,
				indentLeft: 0,
				indentRight: 0,
				indentFirst: 0,
				spaceBefore: 0,
				spaceAfter: 0,
				lineSpacing: { kind: 'multiple', value: 1.2 },
				direction: 'ltr',
				...(tabs ? { tabs } : {}),
			},
		],
		fontFamily: 'Calibri',
		fontSize: 0.14,
		color: '#000000',
		horizontalAlign: 'left',
		verticalAlign: 'top',
		transform: [1, 0, 0, 1, 0, 0],
		width: 6,
		height: 1,
		margins: { left: 0.1, right: 0.1, top: 0.05, bottom: 0.05 },
	}) as VisioText;

describe('tab stops in drawn text', () => {
	it('places each part of a line at the paragraph tab stops', () => {
		const layout = layoutParagraphs(
			text('ab\tcd\tef', [
				{ position: 1, align: 'left' },
				{ position: 3, align: 'right' },
			]),
			measure,
		);
		const parts = layoutTabbedLine(layout.lines[0]!, measure)!;
		expect(parts.map((part) => part.runs.map((item) => item.text).join(''))).toEqual([
			'ab',
			'cd',
			'ef',
		]);
		// The origin is the left margin: stop 1 is 1 in from it; the right stop ends the text.
		expect(parts[0]!.x).toBeCloseTo(0.1);
		expect(parts[1]!.x).toBeCloseTo(1.1);
		expect(parts[2]!.x).toBeCloseTo(3.1 - 0.2);
	});

	it('lays a centred or right-aligned line out from its own start, then moves it (Visio 16)', () => {
		const stops = [{ position: 1, align: 'left' as const }];
		const place = (align: 'center' | 'right') => {
			const layout = layoutParagraphs(text('ab\tcd', stops, align), measure);
			return layoutTabbedLine(layout.lines[0]!, measure)!.map((item) => item.x);
		};
		// The line is 1 in of stop plus "cd" (0.2 in) wide; the stop is 1 in from the line start.
		const centred = place('center');
		expect(centred[0]).toBeCloseTo(3 - 0.6);
		expect(centred[1]! - centred[0]!).toBeCloseTo(1);
		const right = place('right');
		expect(right[1]! + 0.2).toBeCloseTo(5.9);
		expect(right[1]! - right[0]!).toBeCloseTo(1);
		const justified = layoutParagraphs(text('ab\tcd', stops, 'justify'), measure);
		expect(layoutTabbedLine(justified.lines[0]!, measure)).toBeUndefined();
	});

	it('draws one positioned row per part and leaves lines without tabs alone', () => {
		const svg = renderText(text('ab\tcd', [{ position: 2, align: 'left' }]), new Set());
		const rows = Array.from(svg.querySelectorAll('text > tspan'));
		expect(rows).toHaveLength(2);
		expect(rows.map((row) => row.getAttribute('x'))).toEqual(['0.1', '2.1']);
		expect(svg.textContent).not.toContain('\t');
		const plain = renderText(text('ab cd', undefined), new Set());
		expect(plain.querySelectorAll('text > tspan')).toHaveLength(1);
	});
});
