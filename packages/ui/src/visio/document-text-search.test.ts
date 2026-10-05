import { describe, expect, it } from 'vitest';
import type { VisioDocument, VisioShape } from 'ooxml-core/visio';
import { demoDocument } from './demo-document.js';
import {
	EMPTY_TEXT_SEARCH,
	TEXT_SEARCH_LIMITS,
	indexDocumentText,
	searchDocumentText,
	textSearchStatus,
} from './document-text-search.js';

const shape = (id: string, text: string): VisioShape => ({
	...demoDocument.pages[0]!.shapes[0]!,
	id,
	name: id,
	text: { ...demoDocument.pages[0]!.shapes[0]!.text, plainText: text },
	children: [],
});
const documentOf = (...shapes: VisioShape[]): VisioDocument => ({
	...demoDocument,
	pages: [{ ...demoDocument.pages[0]!, shapes }],
});
describe('bounded document text index', () => {
	it('finds literal canonical text once per shape in page and group preorder', () => {
		const group = shape('group', 'Needle needle');
		group.kind = 'group';
		group.children = [shape('child', 'A NEEDLE')];
		const model = documentOf(group, shape('last', 'Needle'));
		model.pages.push({
			...demoDocument.pages[1]!,
			isBackground: true,
			shapes: [shape('child', 'needle')],
		});
		model.pages[0]!.backgroundPageId = model.pages[1]!.id;
		const index = indexDocumentText(model);
		expect(
			searchDocumentText(index, 'NEEDLE').results.map((result) => [result.pageId, result.shapeId]),
		).toEqual([
			['1', 'group'],
			['1', 'child'],
			['1', 'last'],
			['2', 'child'],
		]);
		expect(searchDocumentText(index, 'needle').results).toHaveLength(4);
		expect(searchDocumentText(index, 'Needle needle').results).toHaveLength(1);
	});
	it('excludes hidden subtrees, suppressed group text, names and shape metadata', () => {
		const hidden = shape('hidden', 'needle');
		hidden.hidden = true;
		hidden.children = [shape('hidden-child', 'needle')];
		const group = shape('group', 'needle');
		group.kind = 'group';
		group.groupDisplayMode = 0;
		group.children = [shape('shown-child', 'needle')];
		const metadata = shape('needle', 'other');
		metadata.shapeData = [
			{ id: '0', name: 'needle', valueKind: 'string', value: 'needle', type: 0 },
		];
		const result = searchDocumentText(
			indexDocumentText(documentOf(hidden, group, metadata)),
			'needle',
		);
		expect(result.results.map((result) => result.shapeId)).toEqual(['shown-child']);
	});
	it('does not interpret regex, HTML, Unicode normalization or whitespace patterns', () => {
		const index = indexDocumentText(documentOf(shape('a', '<img onerror="bad"> [A.*]\nÉcole')));
		expect(searchDocumentText(index, '[a.*]').results).toHaveLength(1);
		expect(searchDocumentText(index, '.*').results).toHaveLength(1);
		expect(searchDocumentText(index, '.+').results).toHaveLength(0);
		expect(searchDocumentText(index, 'ÉCOLE').results).toHaveLength(1);
		expect(searchDocumentText(index, 'ecole').results).toHaveLength(0);
		expect(searchDocumentText(index, ']\né').results).toHaveLength(1);
		expect(searchDocumentText(index, '] é').results).toHaveLength(0);
		expect(searchDocumentText(index, '   ').results).toHaveLength(0);
		expect(searchDocumentText(index, '').results).toHaveLength(0);
	});
	it('caps matching shapes, and reports truncation only when another match exists', () => {
		const shapes = Array.from({ length: TEXT_SEARCH_LIMITS.results }, (_, i) =>
			shape(String(i), 'a'),
		);
		const complete = searchDocumentText(indexDocumentText(documentOf(...shapes)), 'a');
		expect(complete.results).toHaveLength(500);
		expect(complete.resultsTruncated).toBe(false);
		const partial = searchDocumentText(
			indexDocumentText(documentOf(...shapes, shape('extra', 'a'))),
			'a',
		);
		expect(partial.results).toHaveLength(500);
		expect(partial.resultsTruncated).toBe(true);
		expect(textSearchStatus(partial)).toBe('500+ matching shapes');
	});
	it('bounds each shape and preview, with an explicit incomplete index', () => {
		const index = indexDocumentText(
			documentOf(shape('a', 'a'.repeat(TEXT_SEARCH_LIMITS.shapeCharacters) + 'end')),
		);
		const result = searchDocumentText(index, 'a');
		expect(index.entries[0]!.foldedText).toHaveLength(TEXT_SEARCH_LIMITS.shapeCharacters);
		expect(result.results[0]!.preview).toHaveLength(TEXT_SEARCH_LIMITS.previewCharacters + 1);
		expect(searchDocumentText(index, 'end').results).toHaveLength(0);
		expect(result.indexTruncated).toBe(true);
		expect(textSearchStatus(result)).toContain('Partial text index');
	});
	it('bounds aggregate copied text, including when no results match', () => {
		const text = 'a'.repeat(TEXT_SEARCH_LIMITS.shapeCharacters);
		const shapes = Array.from({ length: 63 }, (_, i) => shape(String(i), text));
		const index = indexDocumentText(documentOf(...shapes, shape('outside', 'needle')));
		expect(index.entries.reduce((size, entry) => size + entry.foldedText.length, 0)).toBe(
			TEXT_SEARCH_LIMITS.indexedCharacters,
		);
		expect(searchDocumentText(index, 'needle').indexTruncated).toBe(true);
		expect(textSearchStatus(searchDocumentText(index, 'needle'))).toBe(
			'0 matching shapes · Partial text index',
		);
	});
	it('contains cycles and excessive traversal if called outside the validated controller', () => {
		const group = shape('cycle', 'a');
		group.children = [group];
		expect(indexDocumentText(documentOf(group)).truncated).toBe(true);
		const shapes = Array.from({ length: TEXT_SEARCH_LIMITS.shapes + 1 }, (_, i) =>
			shape(String(i), ''),
		);
		expect(indexDocumentText(documentOf(...shapes)).truncated).toBe(true);
		let deep = shape('deep', 'a');
		for (let i = 0; i <= TEXT_SEARCH_LIMITS.depth; i++) {
			const group = shape(String(i), '');
			group.children = [deep];
			deep = group;
		}
		expect(indexDocumentText(documentOf(deep)).truncated).toBe(true);
	});
	it('rejects unbounded queries before reading index entries or converting case', () => {
		const index = {
			get entries(): never {
				throw new Error('Should not scan');
			},
			truncated: false,
		};
		expect(() => searchDocumentText(index, 'a'.repeat(257))).toThrow('at most 256');
		expect(() => searchDocumentText(index, null as unknown as string)).toThrow('must be strings');
	});
	it('freezes state, results and results arrays without mutating snapshots', () => {
		const index = indexDocumentText(documentOf(shape('a', 'a')));
		const result = searchDocumentText(index, 'a');
		expect(Object.isFrozen(result)).toBe(true);
		expect(Object.isFrozen(result.results)).toBe(true);
		expect(Object.isFrozen(result.results[0])).toBe(true);
		expect(Object.isFrozen(EMPTY_TEXT_SEARCH)).toBe(true);
		expect(searchDocumentText(index, 'b').results).toHaveLength(0);
		expect(result.results).toHaveLength(1);
	});
});
