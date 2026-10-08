import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { encodeVisioPlainText } from './plain-text';
import { cell, fixture, row, section, shape } from './test-fixtures';

async function parseText(contents: string, styles = '') {
	const bytes = await fixture({
		pages: [
			{ id: '0', contents: `<Shapes>${shape('1', styles + `<Text>${contents}</Text>`)}</Shapes>` },
		],
	});
	const before = bytes.slice();
	const document = await parseVsdx(bytes);
	expect(bytes).toEqual(before);
	return document.pages[0]!.shapes[0]!.text;
}

describe('native shape text paragraph terminators', () => {
	it.each(['', 'Hello', '\n', '\n\n', 'A\n', 'A\n\n', '😀\nLast\n', ' spaces \t '])(
		'reads logical text and intentional trailing blank paragraphs for %j',
		async (input) => {
			const text = await parseText(encodeVisioPlainText(input));
			expect(text.plainText).toBe(input);
			expect(text.runs.map((run) => run.text).join('')).toBe(input);
			expect(text.paragraphs?.length).toBe(input ? input.split('\n').length : 0);
			for (const paragraph of text.paragraphs ?? []) {
				expect(paragraph.start).toBeLessThanOrEqual(paragraph.end);
				expect(paragraph.end).toBeLessThanOrEqual(input.length);
			}
		},
	);
	it('retains unterminated producer text and removes exactly one stored terminal marker', async () => {
		expect((await parseText('Unterminated')).plainText).toBe('Unterminated');
		expect((await parseText('A\n\n\n')).plainText).toBe('A\n\n');
		expect((await parseText('\n')).plainText).toBe('');
	});
	it('keeps UTF-16 run offsets and the style of a final intentional blank paragraph', async () => {
		const styles =
			section('Character', row(1, '', cell('Font', 'Georgia') + cell('Style', 1))) +
			section('Paragraph', row(1, '', cell('HorzAlign', 2)));
		const text = await parseText('😀\n<pp IX="1"/><cp IX="1"/>\n', styles);
		expect(text.plainText).toBe('😀\n');
		expect(text.runs.map((run) => run.text)).toEqual(['😀\n', '']);
		expect(text.runs[1]).toMatchObject({ bold: true, fontFamily: 'Georgia' });
		expect(text.paragraphs?.map(({ start, end }) => [start, end])).toEqual([
			[0, 2],
			[3, 3],
		]);
		expect(text.paragraphs?.[1]?.horizontalAlign).toBe('right');
	});
	it('does not consume newlines inside field display caches or unknown markup', async () => {
		expect((await parseText('<fld IX="0">Field\n</fld>')).plainText).toBe('Field\n');
		expect((await parseText('A\n<fld IX="0"/>')).plainText).toBe('A\n');
		expect((await parseText('<unknown>Content\n</unknown>')).plainText).toBe('Content\n');
		expect((await parseText('<fld IX="0">Field\n</fld>\n')).plainText).toBe('Field\n');
	});
	it('recognizes CDATA text and preserves nonprinting formatting markers after a terminator', async () => {
		expect((await parseText('<![CDATA[A\n]]><cp IX="0"/><pp IX="0"/>')).plainText).toBe('A');
	});
});
