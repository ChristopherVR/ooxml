import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { normalizeWordClipboardHtml, type WordClipboardList } from './index';
import { clipboardFragment, isWordClipboardHtml } from './word-fragment';

/** HTML Format data Word 16 put on the clipboard (`__fixtures__/clipboard/capture-word-clipboard.ps1`). */
const fixture = (name: string): string =>
	readFileSync(new URL(`../__fixtures__/clipboard/${name}.html`, import.meta.url), 'utf8');

/** What a browser hands a paste handler: the header stripped, from `<html` on. */
const browserForm = (raw: string): string => raw.slice(raw.indexOf('<html'));

function normalize(name: string, mode: 'source' | 'merge' = 'source') {
	const lists: WordClipboardList[] = [];
	const html = normalizeWordClipboardHtml(fixture(name), {
		mode,
		numId: (list) => (lists.push(list), 40 + lists.length),
		paragraphStyleId: (style) => style.replace(/\s/g, '').replace(/^heading/, 'Heading'),
		characterStyleId: (style) => style.replace(/\s/g, ''),
	});
	return { html: html ?? '', lists };
}

describe('Word clipboard HTML', () => {
	it('recognises Word HTML and leaves other HTML alone', () => {
		expect(isWordClipboardHtml(fixture('word-styles'))).toBe(true);
		expect(normalizeWordClipboardHtml('<p>Plain <b>web</b> text</p>')).toBeNull();
	});

	it('finds the copied fragment in the raw and the browser form', () => {
		const raw = fixture('word-bullets');
		expect(raw.startsWith('Version:')).toBe(true);
		for (const html of [raw, browserForm(raw)]) {
			const fragment = clipboardFragment(html);
			expect(fragment).toContain('Before list');
			expect(fragment).not.toContain('<style');
		}
		expect(normalizeWordClipboardHtml(browserForm(raw))).toBe(normalizeWordClipboardHtml(raw));
	});

	it('turns mso-list paragraphs into list paragraphs with levels', () => {
		const { html, lists } = normalize('word-numbered');
		expect(lists).toHaveLength(1);
		expect(lists[0]!.key).toBe('l0:lfo1');
		const [one, two, three] = lists[0]!.levels;
		expect(one).toMatchObject({ level: 0, numFmt: 'decimal', lvlText: '%1)', start: 1 });
		expect(one).toMatchObject({ indentLeftTwips: 360, hangingTwips: 360 });
		expect(two).toMatchObject({ level: 1, numFmt: 'lowerLetter', lvlText: '%2)' });
		expect(three).toMatchObject({ level: 2, numFmt: 'lowerRoman', lvlText: '%3)' });
		expect(lists[0]!.levels).toHaveLength(9);
		// The marker Word wrote for browsers is gone; the list draws its own.
		expect(html).not.toMatch(/1\)|a\)|supportLists|mso-/);
		expect(html.match(/data-ilvl="(\d)"/g)).toEqual([
			'data-ilvl="0"',
			'data-ilvl="1"',
			'data-ilvl="2"',
			'data-ilvl="1"',
			'data-ilvl="0"',
		]);
		expect(html).toContain(
			'<p data-docx-paragraph-style="ListParagraph" data-num-id="41" data-ilvl="0">Step one</p>',
		);
	});

	it('reads bullet glyphs and fonts, with the indent the style supplies', () => {
		const { html, lists } = normalize('word-bullets');
		const [one, two, three] = lists[0]!.levels;
		expect(one).toMatchObject({
			numFmt: 'bullet',
			lvlText: String.fromCodePoint(0xf0b7),
			markerFormat: { fontFamily: 'Symbol' },
			// Level 1 sets no margin: the List Paragraph style's half inch applies.
			indentLeftTwips: 720,
			hangingTwips: 360,
		});
		expect(two).toMatchObject({ lvlText: 'o', markerFormat: { fontFamily: 'Courier New' } });
		expect(three).toMatchObject({
			lvlText: String.fromCodePoint(0xf0a7),
			markerFormat: { fontFamily: 'Wingdings' },
		});
		expect(html).toContain('<p data-docx-paragraph-style="Normal">Before list</p>');
		expect(html).toContain('data-num-id="41" data-ilvl="2">Deeper bullet</p>');
		expect(html).not.toContain('·');
	});

	it('infers levels from the markers when the stylesheet is missing', () => {
		const fragment = clipboardFragment(fixture('word-numbered'));
		const lists: WordClipboardList[] = [];
		const html = normalizeWordClipboardHtml(
			`<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>${fragment}</body></html>`,
			{ numId: (list) => (lists.push(list), 7) },
		);
		expect(lists[0]!.levels.map(({ level, numFmt, lvlText }) => [level, numFmt, lvlText])).toEqual([
			[0, 'decimal', '%1)'],
			[1, 'lowerLetter', '%2)'],
			[2, 'lowerRoman', '%3)'],
		]);
		expect(html).toContain('data-num-id="7" data-ilvl="2">Sub sub step</p>');
		const bullets = normalizeWordClipboardHtml(
			`<html xmlns:w="urn:schemas-microsoft-com:office:word"><body>${clipboardFragment(fixture('word-bullets'))}</body></html>`,
			{ numId: (list) => (lists.push(list), 8) },
		);
		expect(bullets).toContain('data-num-id="8"');
		expect(lists[1]!.levels[0]).toMatchObject({
			numFmt: 'bullet',
			lvlText: String.fromCodePoint(0xf0b7),
			markerFormat: { fontFamily: 'Symbol' },
		});
	});

	it('reads the form a sanitising clipboard read gives (markers as comments, no stylesheet)', () => {
		const sanitised = clipboardFragment(fixture('word-numbered'))
			.replace(/<!\[if ([^\]]*)\]>/g, '<!--[if $1]-->')
			.replace(/<!\[endif\]>/g, '<!--[endif]-->');
		const lists: WordClipboardList[] = [];
		const html = normalizeWordClipboardHtml(`<html><body>${sanitised}</body></html>`, {
			numId: (list) => (lists.push(list), 3),
		})!;
		expect(lists[0]!.levels.map((level) => level.numFmt)).toEqual([
			'decimal',
			'lowerLetter',
			'lowerRoman',
		]);
		expect(html).toContain('<p data-num-id="3" data-ilvl="1">Sub step a</p>');
		expect(html).not.toMatch(/1\)|supportLists/);
	});

	it('maps classes and headings to style names', () => {
		const { html } = normalize('word-styles');
		expect(html).toContain('<p data-docx-paragraph-style="Heading1">Clipboard Heading</p>');
		expect(html).toContain('<p data-docx-paragraph-style="Quote">A quoted paragraph.</p>');
		expect(html).toContain('<p data-docx-paragraph-style="Heading2">Second level heading</p>');
		// A custom style's class is its name without spaces; the stylesheet says its real name.
		expect(html).toContain('<p data-docx-paragraph-style="CalloutText">Custom styled</p>');
		expect(html).not.toMatch(/<h\d|o:p|class=/);
	});

	it('keeps Word-only run formatting as run properties and drops mso- CSS', () => {
		const { html } = normalize('word-styles');
		expect(html).toContain('<b>bold</b>');
		expect(html).toContain('<span data-run-props="{&quot;smallCaps&quot;:true}">small caps</span>');
		expect(html).toContain('<span style="background-color:#ffff00">highlighted');
		expect(html).toContain(
			'<span data-run-props="{&quot;doubleStrike&quot;:true}">double struck</span>',
		);
		expect(html).toContain(
			'<u><span data-run-props="{&quot;underlineStyle&quot;:&quot;double&quot;}">double underline</span></u>',
		);
		expect(html).toContain('<span style="color:#ff0000">red</span>');
		expect(html).not.toContain('mso-');
	});

	it('keeps fields with their codes and notes as text', () => {
		const { html } = normalize('word-table-notes');
		expect(html).toContain('Page <span data-field=" PAGE " data-field-simple="1">1</span></p>');
		expect(html).toContain('Text with a note<sup>1</sup></p>');
		expect(html).toMatch(/<sup>1<\/sup> The\s+footnote text\./);
		expect(html).not.toMatch(/_ftn|<hr|supportFootnotes|\[1\]/);
		expect(html).toMatch(/<table>\s*<tr>\s*<td>/);
		expect(html).toContain('<p data-docx-paragraph-style="Normal"><b>A1</b></p>');
		// Word's empty paragraph after a table (`<o:p>&nbsp;</o:p>`) stays an empty paragraph.
		expect(html).toContain('<p data-docx-paragraph-style="Normal"></p>');
	});

	it('merge formatting keeps lists and emphasis but not styles or fonts', () => {
		const styles = normalize('word-styles', 'merge').html;
		expect(styles).toContain('<p>Clipboard Heading</p>');
		expect(styles).toContain('<b>bold</b>');
		expect(styles).toContain('<u>double underline</u>');
		expect(styles).not.toMatch(/data-run-props|style=|data-docx/);
		const numbered = normalize('word-numbered', 'merge').html;
		expect(numbered).toContain('<p data-num-id="41" data-ilvl="0">Step one</p>');
	});
});
