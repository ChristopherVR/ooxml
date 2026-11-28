import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createEstimatedMeasurer } from '../pdf-export/font-measurer';
import { layoutDocumentModel } from './index';
import { justifiedSqueezeBudgetPx } from './paragraph-layout';

// Issue #381: Word's squeeze budget for the last word of a justified line is
// min(0.25 of the line's spaces, (word + preceding space) * (0.3325 + 0.25 / N)), N the spaces on
// the line before the word. The fixtures are the reporter's Word 16 files; each group is a heading
// paragraph, then paragraphs whose last word on the first line needs 0.05 pt more squeeze each.
// The counts are the paragraphs whose first line keeps that word in Word's own PDF. Only Arial and
// Calibri are checked: their advances are bundled (`text/font-metrics`), Century Gothic's are not.

/** First line text of each paragraph, in document order, with the heading paragraphs. */
async function firstLines(name: string): Promise<string[]> {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-parity/issue-381-${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, createEstimatedMeasurer());
	const seen = new Set<string>();
	const lines: string[] = [];
	for (const page of pages)
		for (const column of page.columns)
			for (const block of column.blocks) {
				if (block.kind !== 'paragraph' || seen.has(block.blockId)) continue;
				seen.add(block.blockId);
				const text = (block.lines[0]?.fragments ?? []).map((fragment) => fragment.text).join('');
				lines.push(text.replace(/\s+/g, ' ').trim());
			}
	return lines;
}

/** For each heading naming the last word, how many of its paragraphs keep that word. */
async function keptByGroup(name: string): Promise<Map<string, number>> {
	const kept = new Map<string, number>();
	let group: { key: string; word: string } | undefined;
	for (const line of await firstLines(name)) {
		const heading = /^(\w[\w ]*? \d+ pt), (\d+) spaces before the last word "(\w+)"/.exec(line);
		if (heading) {
			group = { key: `${heading[1]} ${heading[2]} ${heading[3]}`, word: heading[3] ?? '' };
			kept.set(group.key, 0);
		} else if (group && line.startsWith('Lorem') && line.split(' ').at(-1) === group.word)
			kept.set(group.key, (kept.get(group.key) ?? 0) + 1);
	}
	return kept;
}

describe('justified squeeze budget scales with the spaces on the line (issue #381)', () => {
	it('keeps "i" after 3 and 4 Arial spaces and wraps "aliquet" after 20 Calibri ones', async () => {
		const kept = await keptByGroup('justify-word-budget-by-spaces');
		expect(kept.get('Arial 12 pt 3 i')).toBe(1);
		expect(kept.get('Arial 12 pt 4 i')).toBe(1);
		expect(kept.get('Calibri 11 pt 20 aliquet')).toBe(0);
	});

	it.each([
		{ name: 'sweep-few-spaces', key: 'Arial 12 pt 3 i', count: 9 },
		{ name: 'sweep-few-spaces', key: 'Arial 12 pt 4 i', count: 6 },
		{ name: 'sweep-few-spaces', key: 'Arial 12 pt 5 i', count: 4 },
		{ name: 'sweep-short-word-by-spaces', key: 'Calibri 11 pt 7 a', count: 5 },
		{ name: 'sweep-short-word-by-spaces', key: 'Calibri 11 pt 12 a', count: 5 },
		{ name: 'sweep-short-word-by-spaces', key: 'Calibri 11 pt 20 a', count: 5 },
		{ name: 'sweep-long-word-many-spaces', key: 'Calibri 11 pt 20 quam', count: 5 },
		{ name: 'sweep-long-word-many-spaces', key: 'Calibri 11 pt 20 aliquet', count: 5 },
	])(
		'$name: $key keeps the word in $count paragraphs as Word does',
		async ({ name, key, count }) => {
			expect((await keptByGroup(name)).get(key)).toBe(count);
		},
	);

	it('takes the word part from the spaces on the line', () => {
		// "a" plus its space, 10 px, on lines of 7 and 20 spaces of 5 px each.
		expect(justifiedSqueezeBudgetPx(35, 7, 10)).toBeCloseTo(10 * (0.3325 + 0.25 / 7));
		expect(justifiedSqueezeBudgetPx(100, 20, 10)).toBeCloseTo(3.45);
		// A long word: the quarter of the spaces decides.
		expect(justifiedSqueezeBudgetPx(15, 3, 100)).toBeCloseTo(3.75);
		expect(justifiedSqueezeBudgetPx(0, 0, 10)).toBe(0);
	});
});
