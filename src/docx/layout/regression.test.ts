import { describe, expect, it } from 'vitest';
import {
	createDocument,
	signedTwips,
	twips,
	type DocumentModel,
	type Paragraph,
	type SectionProperties,
	type Table,
} from '../index.js';
import { layoutDocumentModel } from './layout.js';
import { createFakeMeasurer } from './measure.js';
import { at } from './test-helpers.js';

const measurer = createFakeMeasurer();
const paragraph = (id: string, text: string, extra: Partial<Paragraph> = {}): Paragraph => ({
	type: 'paragraph',
	id,
	runs: [{ text }],
	...extra,
});
function model(blocks: DocumentModel['blocks'], height = 300): DocumentModel {
	const doc = createDocument();
	doc.blocks = blocks;
	doc.page = { ...doc.page, height, marginTop: 20, marginBottom: 20 };
	return doc;
}
const blockIds = (page: ReturnType<typeof layoutDocumentModel>['pages'][number]) =>
	page.columns.flatMap((column) => column.blocks.map((block) => block.blockId));

describe('pagination determinism', () => {
	const blocks = Array.from({ length: 60 }, (_, i) =>
		paragraph(`p${i}`, `Paragraph ${i} `.repeat(12)),
	);

	it('returns structurally identical results for the same model and measurer', () => {
		const first = layoutDocumentModel(model(blocks), measurer);
		const second = layoutDocumentModel(model(blocks), measurer);
		expect(first.pages.length).toBeGreaterThan(3);
		expect(JSON.stringify(second)).toBe(JSON.stringify(first));
	});

	it('does not mutate the model it lays out', () => {
		const doc = model(blocks);
		const before = JSON.stringify(doc);
		layoutDocumentModel(doc, measurer);
		expect(JSON.stringify(doc)).toBe(before);
	});

	it('places every block exactly once and in document order', () => {
		const result = layoutDocumentModel(model(blocks), measurer);
		const placed = result.pages.flatMap(blockIds);
		const unique = [...new Set(placed)];
		expect(unique).toEqual(blocks.map((block) => block.id));
		expect(result.pages.map((page) => page.index)).toEqual(result.pages.map((_, i) => i));
	});
});

describe('page breaks from the document model', () => {
	it('honours an explicit page-break run', () => {
		const result = layoutDocumentModel(
			model([
				{ type: 'paragraph', id: 'a', runs: [{ text: 'one' }, { text: '', break: 'page' }] },
				paragraph('b', 'two'),
			]),
			measurer,
		);
		expect(result.pages).toHaveLength(2);
		expect(blockIds(at(result.pages, 1))).toEqual(['b']);
	});

	it('honours pageBreakBefore', () => {
		const result = layoutDocumentModel(
			model([paragraph('a', 'one'), paragraph('b', 'two', { pageBreakBefore: true })]),
			measurer,
		);
		expect(result.pages.map(blockIds)).toEqual([['a'], ['b']]);
	});

	it('does not add a blank page when the break is the last thing in the document', () => {
		const result = layoutDocumentModel(
			model([paragraph('a', 'one', { runs: [{ text: 'one' }, { text: '', break: 'page' }] })]),
			measurer,
		);
		const placed = result.pages.flatMap(blockIds);
		expect(placed).toEqual(['a']);
	});
});

describe('tables splitting across pages', () => {
	const table = (rows: number): Table => ({
		type: 'table',
		id: 't',
		rows: Array.from({ length: rows }, (_, i) => [
			{ paragraphs: [paragraph(`c${i}`, `Row ${i}`)] },
		]),
	});
	const placedRows = (result: ReturnType<typeof layoutDocumentModel>) =>
		result.pages.flatMap((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) =>
					block.kind === 'table' ? block.rows.filter((row) => !row.repeated) : [],
				),
			),
		);

	it('splits a long table at row boundaries without losing or repeating rows', () => {
		const result = layoutDocumentModel(model([table(40)], 200), measurer);
		const texts = placedRows(result).map((row) => at(at(row.cells, 0), 0).lines.length);
		expect(result.pages.length).toBeGreaterThan(1);
		expect(texts).toHaveLength(40);
	});

	it('keeps every placed table inside the printable height of its page', () => {
		const result = layoutDocumentModel(model([table(40)], 200), measurer);
		for (const page of result.pages)
			for (const block of page.columns.flatMap((column) => column.blocks))
				expect(block.yPx + block.heightPx).toBeLessThanOrEqual(
					page.heightPx - page.marginBottomPx + 0.5,
				);
	});
});

describe('sections and the header/footer contract', () => {
	const section = (endsAtBlockId: string, type: 'nextPage' | 'continuous'): SectionProperties => ({
		endsAtBlockId,
		type,
		pageWidthTwips: twips(12240),
		pageHeightTwips: twips(3000),
		orientation: 'portrait',
		marginTopTwips: signedTwips(300),
		marginRightTwips: twips(1440),
		marginBottomTwips: signedTwips(300),
		marginLeftTwips: twips(1440),
		columns: { count: 1, equalWidth: true },
		titlePage: true,
	});

	it('reports section index and page-in-section on every page, for first-page headers', () => {
		const doc = model([paragraph('a', 'x '.repeat(500)), paragraph('b', 'y ')]);
		doc.sections = [section('a', 'nextPage'), section('b', 'nextPage')];
		const result = layoutDocumentModel(doc, measurer);
		const first = result.pages.filter((page) => page.sectionIndex === 0);
		const second = result.pages.filter((page) => page.sectionIndex === 1);
		expect(first.length).toBeGreaterThan(1);
		expect(first.map((page) => page.pageInSection)).toEqual(first.map((_, i) => i));
		expect(second.map((page) => page.pageInSection)).toEqual([0]);
		expect(result.pages.map((page) => page.index)).toEqual(result.pages.map((_, i) => i));
	});

	it('does not lay out header or footer content itself (hosts render it from the model)', () => {
		const result = layoutDocumentModel(model([paragraph('a', 'body')]), measurer);
		expect(result.pages.flatMap(blockIds)).toEqual(['a']);
	});
});
