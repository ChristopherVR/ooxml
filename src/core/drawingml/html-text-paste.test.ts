import { describe, expect, it } from 'vitest';

import { drawingParagraphsText, htmlToDrawingParagraphs } from './html-text-paste';
import type { DrawingParagraph } from './text-model';

const runs = (p: DrawingParagraph | undefined) =>
	(p?.items ?? []).flatMap((item) => (item.kind === 'run' ? [item] : []));

describe('htmlToDrawingParagraphs', () => {
	it('splits blocks into paragraphs and keeps line breaks inside one', () => {
		const { paragraphs } = htmlToDrawingParagraphs('<p>One<br>two</p>\n<div>Three</div>');
		expect(paragraphs).toHaveLength(2);
		expect(paragraphs[0]?.items.map((item) => item.kind)).toEqual(['run', 'break', 'run']);
		expect(drawingParagraphsText(paragraphs)).toBe('One\ntwo\nThree');
	});

	it('reads bold, italic, underline, strike, super and subscript', () => {
		const { paragraphs } = htmlToDrawingParagraphs(
			'<p><b>B</b><i>I</i><u>U</u><s>S</s>x<sup>2</sup>H<sub>2</sub></p>',
		);
		const props = runs(paragraphs[0]).map((run) => run.properties);
		expect(props[0]).toEqual({ bold: true });
		expect(props[1]).toEqual({ italic: true });
		expect(props[2]).toEqual({ underline: 'sng' });
		expect(props[3]).toEqual({ strike: 'sngStrike' });
		expect(props[5]).toEqual({ baseline: 0.3 });
		expect(props[7]).toEqual({ baseline: -0.25 });
	});

	it('reads font family, size, colour and highlight from CSS and <font>', () => {
		const { paragraphs } = htmlToDrawingParagraphs(
			"<p><span style=\"font-family:'Segoe UI', sans-serif; font-size:18pt; color:#c00; " +
				'background-color: rgb(255, 255, 0)">A</span><font face="Arial" color="blue" size="5">B</font></p>',
		);
		const [a, b] = runs(paragraphs[0]);
		expect(a?.properties).toMatchObject({
			latin: { typeface: 'Segoe UI' },
			sizePt: 18,
			fill: { kind: 'solid', color: { kind: 'srgb', value: 'CC0000' } },
			highlight: { kind: 'srgb', value: 'FFFF00' },
		});
		expect(b?.properties).toMatchObject({
			latin: { typeface: 'Arial' },
			sizePt: 18,
			fill: { kind: 'solid', color: { value: '0000FF' } },
		});
	});

	it('converts px sizes and paragraph alignment', () => {
		const { paragraphs } = htmlToDrawingParagraphs(
			'<p style="text-align:center"><span style="font-size:16px">c</span></p><p align="right">r</p>',
		);
		expect(paragraphs[0]?.properties?.align).toBe('ctr');
		expect(runs(paragraphs[0])[0]?.properties?.sizePt).toBe(12);
		expect(paragraphs[1]?.properties?.align).toBe('r');
	});

	it('turns lists into bulleted and numbered paragraphs with levels', () => {
		const { paragraphs } = htmlToDrawingParagraphs(
			'<ul><li>a<ul><li>b</li></ul></li><li>c</li></ul><ol type="A" start="3"><li>d</li></ol>',
		);
		expect(drawingParagraphsText(paragraphs)).toBe('a\nb\nc\nd');
		expect(paragraphs[0]?.properties?.bullet).toEqual({ kind: 'char', char: '\u2022' });
		expect(paragraphs[1]?.properties?.level).toBe(1);
		expect(paragraphs[2]?.properties?.level).toBeUndefined();
		expect(paragraphs[3]?.properties?.bullet).toEqual({
			kind: 'autoNumber',
			scheme: 'alphaUcPeriod',
			startAt: 3,
		});
	});

	it("reads Word's typed list markers as bullets instead of text", () => {
		const html =
			"<p class=MsoListParagraph style='mso-list:l0 level2 lfo1'><![if !supportLists]>" +
			"<span>2.<span style='font:7.0pt'>&nbsp;&nbsp;</span></span><![endif]>Second</p>" +
			"<p style='mso-list:l1 level1 lfo2'><span style='mso-list:Ignore'>\u00b7<span>&nbsp;</span></span>Dot</p>";
		const { paragraphs } = htmlToDrawingParagraphs(html);
		expect(drawingParagraphsText(paragraphs)).toBe('Second\nDot');
		expect(paragraphs[0]?.properties).toMatchObject({
			level: 1,
			bullet: { kind: 'autoNumber', scheme: 'arabicPeriod', startAt: 2 },
		});
		expect(paragraphs[1]?.properties?.bullet).toEqual({ kind: 'char', char: '\u2022' });
	});

	it('keeps safe links as placeholder relationships and drops script links', () => {
		const { paragraphs, hyperlinks } = htmlToDrawingParagraphs(
			'<p><a href="https://example.com/x" title="Tip">ok</a> <a href="javascript:alert(1)">bad</a> ' +
				'<a href=" jav&#x09;ascript:alert(1)">worse</a></p>',
		);
		const [ok, ...rest] = runs(paragraphs[0]);
		expect(ok?.properties?.hyperlinkClick).toEqual({ relId: 'htmlLink1', tooltip: 'Tip' });
		expect(hyperlinks).toEqual({ htmlLink1: 'https://example.com/x' });
		expect(rest.every((run) => !run.properties?.hyperlinkClick)).toBe(true);
		expect(drawingParagraphsText(paragraphs)).toBe('ok bad worse');
	});

	it('slices the CF_HTML fragment and drops head, style, script and comments', () => {
		const html =
			'Version:0.9\r\nStartHTML:0000000105\r\n<html><head><style>p{color:red}</style>' +
			'<title>T</title></head><body><!--StartFragment--><p>Kept<script>alert(1)</script>' +
			'<!-- hidden --></p><!--EndFragment--><p>outside</p></body></html>';
		expect(drawingParagraphsText(htmlToDrawingParagraphs(html).paragraphs)).toBe('Kept');
	});

	it('collapses whitespace, keeps <pre> text and degrades unknown markup to text', () => {
		const { paragraphs } = htmlToDrawingParagraphs(
			'<p>  a \n  b  </p><pre>x  y\nz</pre><custom-tag data-x="1">c</custom-tag> 1 < 2',
		);
		expect(drawingParagraphsText(paragraphs)).toBe('a b\nx  y\nz\nc 1 < 2');
	});

	it('decodes entities safely, including out-of-range references', () => {
		const { paragraphs } = htmlToDrawingParagraphs(
			'<p>&amp;&lt;&#65;&#x1F600;&#x110000;&#0;&bogus;</p>',
		);
		expect(drawingParagraphsText(paragraphs)).toBe('&<A\u{1F600}\ufffd\ufffd&bogus;');
	});
});

describe('htmlToDrawingParagraphs limits', () => {
	it('caps the input length', () => {
		const result = htmlToDrawingParagraphs(`<p>${'a'.repeat(5000)}</p>`, { maxInputLength: 100 });
		expect(result.truncated).toBe(true);
		expect(drawingParagraphsText(result.paragraphs).length).toBeLessThanOrEqual(100);
	});

	it('caps the text length', () => {
		const result = htmlToDrawingParagraphs(`<p>${'a'.repeat(500)}</p>`, { maxTextLength: 10 });
		expect(result.truncated).toBe(true);
		expect(drawingParagraphsText(result.paragraphs)).toBe('a'.repeat(10));
	});

	it('caps the paragraph count', () => {
		const result = htmlToDrawingParagraphs('<p>x</p>'.repeat(50), { maxParagraphs: 5 });
		expect(result.paragraphs).toHaveLength(5);
		expect(result.truncated).toBe(true);
	});

	it('keeps text but not formatting beyond the depth cap', () => {
		const html = `${'<b><i>'.repeat(5000)}deep${'</i></b>'.repeat(5000)}`;
		const result = htmlToDrawingParagraphs(html, { maxDepth: 16 });
		expect(result.truncated).toBe(true);
		expect(drawingParagraphsText(result.paragraphs)).toBe('deep');
	});

	it('caps the token count and reads the rest as text', () => {
		const result = htmlToDrawingParagraphs(`${'<span>a</span>'.repeat(100)}`, { maxTokens: 20 });
		expect(result.truncated).toBe(true);
		expect(drawingParagraphsText(result.paragraphs)).toBe('a'.repeat(100));
	});

	it('stays linear on hostile markup', () => {
		const hostile = [
			'<a "'.repeat(20000),
			'<!--'.repeat(20000),
			'<style>'.repeat(20000),
			'<'.repeat(200000),
			'<![if '.repeat(20000),
		];
		for (const html of hostile) {
			const started = Date.now();
			htmlToDrawingParagraphs(html);
			expect(Date.now() - started).toBeLessThan(2000);
		}
	});
});
