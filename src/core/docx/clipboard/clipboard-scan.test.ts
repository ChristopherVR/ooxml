import { describe, expect, it } from 'vitest';

import { readStyleRules } from './css';
import { clipboardFragment, isWordClipboardHtml } from './word-fragment';

describe('clipboard HTML scanning', () => {
	it('reads style blocks whose closing tag carries whitespace, and unclosed ones', () => {
		const rules = readStyleRules(
			'<style><!-- p.MsoNormal { margin: 0 } /* note */ --></style ><style>h1 { color: red }',
		);
		expect(rules.map((rule) => rule.selector)).toEqual(['p.MsoNormal', 'h1']);
		expect(rules[1]?.declarations.get('color')).toBe('red');
	});

	it('drops an unclosed css comment to the end of the block', () => {
		expect(readStyleRules('<style>a { b: c } /* open p { d: e }</style>')).toHaveLength(1);
	});

	it('finds Word in a meta tag and the body without a closing tag', () => {
		expect(isWordClipboardHtml('<meta name=Generator content="Microsoft Word 15">')).toBe(true);
		expect(isWordClipboardHtml("<META content='Word.Document' name=ProgId>")).toBe(true);
		expect(isWordClipboardHtml('<meta content="Something else">')).toBe(false);
		expect(clipboardFragment('<html><BODY lang=EN>text')).toBe('text');
		expect(clipboardFragment('<bodyx>no</bodyx>')).toBe('<bodyx>no</bodyx>');
	});

	it('stays linear on hostile input', () => {
		const hostile = '<meta'.repeat(50_000) + '<style'.repeat(50_000) + '<body'.repeat(50_000);
		const started = performance.now();
		isWordClipboardHtml(hostile);
		clipboardFragment(hostile);
		readStyleRules(hostile);
		expect(performance.now() - started).toBeLessThan(2000);
	});
});
