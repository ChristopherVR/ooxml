import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx, type HeaderFooterContent } from '../index';
import { layoutDocumentModel } from './layout';
import { createFakeMeasurer } from './measure';
import {
	headerFooterForPage,
	pageNumbers,
	pageNumberValues,
	sectionPageCounts,
	fieldDisplayText,
} from './page-fields';

const fixture = (name: string) =>
	new URL(`./fixtures/continuous-page-fields/${name}`, import.meta.url);
const evidence = JSON.parse(await readFile(fixture('evidence.json'), 'utf8')) as {
	cases: { name: string; pages: { header: string; footer: string }[] }[];
};
function text(
	content: HeaderFooterContent | undefined,
	page: string,
	sectionPages: number,
	numPages: number,
): string {
	return content!.blocks
		.map((block) =>
			block.type === 'paragraph'
				? block.runs
						.map((run) => {
							return fieldDisplayText(run, {
								page,
								sectionPages: String(sectionPages),
								numPages: String(numPages),
							});
						})
						.join('')
				: '',
		)
		.join('');
}

describe('native Word page-field references', () => {
	it.each(evidence.cases)('matches visible headers and page fields in $name', async (reference) => {
		const loaded = await loadDocx(
			new Uint8Array(await readFile(fixture(`${reference.name}.docx`))),
		);
		const result = layoutDocumentModel(loaded.model, createFakeMeasurer());
		const labels = pageNumbers(loaded.model, result.pages);
		const values = pageNumberValues(loaded.model, result.pages);
		const counts = sectionPageCounts(result.pages);
		expect(result.pages).toHaveLength(reference.pages.length);
		const actual = result.pages.map((page, index) => ({
			header: text(
				headerFooterForPage(loaded.model, page, values[index]!, 'headers'),
				labels[index]!,
				counts.get(page.sectionIndex)!,
				result.pages.length,
			),
			footer: text(
				headerFooterForPage(loaded.model, page, values[index]!, 'footers'),
				labels[index]!,
				counts.get(page.sectionIndex)!,
				result.pages.length,
			),
		}));
		expect(actual).toEqual(reference.pages);
	});
});
