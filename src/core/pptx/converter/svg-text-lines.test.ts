import { describe, expect, it, vi } from 'vitest';
import { SvgExporter } from './SvgExporter';
import { svgTextLines } from './svg-text-lines';

describe('static SVG text wrapping', () => {
	it('uses host font metrics and falls back when they are unavailable', () => {
		const measure = vi.fn(() => 60);
		const segments = [{ text: 'one two', style: { bold: true, fontFamily: 'Verdana' } }];
		expect(svgTextLines(segments, 100, 24, 'Arial', measure).map((run) => run.text)).toEqual([
			'one',
			'two',
		]);
		expect(measure).toHaveBeenCalledWith('one', {
			family: 'Verdana',
			size: 24,
			bold: true,
			italic: false,
		});
		expect(svgTextLines(segments, 100, 24, 'Arial', () => Number.NaN)).toEqual(
			svgTextLines(segments, 100, 24, 'Arial'),
		);
	});
	it('wraps rich text while preserving formatting across line breaks', () => {
		const runs = svgTextLines(
			[{ text: 'Alpha beta gamma delta', style: { bold: true } }],
			100,
			24,
			'Arial',
		);
		expect(runs.filter((run) => run.lineStart).length).toBeGreaterThan(1);
		expect(runs.map((run) => run.text).join(' ')).toBe('Alpha beta gamma delta');
		expect(runs.every((run) => run.style.bold)).toBe(true);
	});
	it('keeps a word split across formatting runs together and honors explicit breaks', () => {
		const runs = svgTextLines(
			[
				{ text: 'Hel', style: { bold: true } },
				{ text: 'lo\nWorld', style: {} },
				{ text: '', style: {}, isParagraphBreak: true },
				{ text: 'Again', style: {} },
			],
			50,
			24,
			'Arial',
		);
		expect(runs.map((run) => [run.text, run.lineStart])).toEqual([
			['Hel', true],
			['lo', false],
			['World', true],
			['Again', true],
		]);
	});
	it('exports wrapped and explicitly unwrapped text without losing XML escaping', () => {
		const element = {
			type: 'text' as const,
			id: 'text',
			x: 0,
			y: 0,
			width: 110,
			height: 200,
			text: 'Alpha beta gamma & delta',
			textStyle: { fontSize: 24 },
		};
		const slide = { id: 'slide', rId: 'rId1', slideNumber: 1, elements: [element] };
		const wrapped = SvgExporter.exportSlide(slide, 960, 540);
		expect(wrapped.match(/<tspan/g)!.length).toBeGreaterThan(1);
		expect(wrapped).toContain('&amp;');
		element.textStyle = { fontSize: 24, textWrap: 'none' } as typeof element.textStyle;
		expect(SvgExporter.exportSlide(slide, 960, 540).match(/<tspan/g)).toHaveLength(1);
	});
});
