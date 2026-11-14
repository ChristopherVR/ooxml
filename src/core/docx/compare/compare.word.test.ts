import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../parse';
import { compareDocuments } from './compare';
import { revisionSpans } from './spans';
import type { CompareGranularity } from './types';

// Pairs built with Word 16 and Word's own Review > Compare result for each, saved by
// __fixtures__/compare/generate-word-compare.ps1 (Application.CompareDocuments).
const FIXTURES = join(__dirname, '..', '__fixtures__', 'compare');
const fixture = (name: string) => readFile(join(FIXTURES, name));

const CASES: [string, CompareGranularity][] = [
	['words', 'word'],
	['format', 'word'],
	['tracked', 'word'],
	['chars', 'character'],
	['table', 'word'],
];

describe('compareDocuments against Word 16 Compare Documents', () => {
	for (const [name, granularity] of CASES)
		it(`marks the same spans as Word: ${name}`, async () => {
			const [original, revised, word] = await Promise.all([
				fixture(`${name}-original.docx`),
				fixture(`${name}-revised.docx`),
				fixture(`${name}-word.docx`),
			]);
			const result = await compareDocuments(original, revised, {
				author: 'Compare Author',
				date: '2026-01-01T00:00:00Z',
				granularity,
			});
			const expected = revisionSpans((await loadDocx(word)).model);
			expect(revisionSpans(result.model)).toEqual(expected);
		});
});
