import { describe, expect, it } from 'vitest';
import type { VisioDocument, VisioShape } from '../model';
import { demoDocument } from './demo-document';
import {
	visioTextReplaceCommands,
	visioTextReplaceFindNext,
	visioTextOccurrenceSelection,
	visioTextReplaceOccurrences,
	visioTextReplacePlan,
	type VisioTextReplaceRequest,
	type VisioTextReplaceEdit,
	type VisioTextReplacePlanEdit,
} from './text-replace';

const request: VisioTextReplaceRequest = {
	query: 'a',
	replacement: 'X',
	matchCase: true,
	scope: 'all-pages',
	pageId: '0',
	mode: 'all',
};
function shape(id: string, text: string, children: VisioShape[] = []): VisioShape {
	const result = structuredClone(demoDocument.pages[0]!.shapes[0]!);
	result.id = id;
	result.text.plainText = text;
	result.children = children;
	return result;
}
function model(...texts: string[]): VisioDocument {
	const result = structuredClone(demoDocument);
	result.pages = [
		{
			...result.pages[0]!,
			id: '0',
			shapes: texts.map((text, index) => shape(String(index + 1), text)),
		},
	];
	return result;
}

const rangeEdit = (edit: Readonly<VisioTextReplaceEdit>) => {
	if (edit.type !== 'replace-text-ranges') throw new Error('Expected a range command');
	return edit;
};
describe('literal full-text replacement planning', () => {
	it('navigates complete occurrences and wraps without using lowercase Find previews', () => {
		const source = model('İ a a');
		const first = visioTextReplaceFindNext(source, request)!;
		expect(first.start).toBe(2);
		const second = visioTextReplaceFindNext(source, request, first)!;
		expect(second.start).toBe(4);
		expect(visioTextReplaceFindNext(source, request, second)).toEqual(first);
		expect(() => visioTextReplaceFindNext(source, request, { ...first, start: 3, end: 4 })).toThrow(
			/stale/,
		);
	});
	it('owns page-qualified nested group navigation records', () => {
		const source = model('a');
		source.pages[0]!.shapes[0]!.children = [shape('2', '😀a')];
		const occurrence = { pageId: '0', shapeId: '2', start: 2, end: 3 };
		const target = visioTextOccurrenceSelection(source, occurrence)!;
		expect(target).toEqual({
			pageIndex: 0,
			selection: { id: '2', name: source.pages[0]!.shapes[0]!.children[0]!.name, pageId: '0' },
		});
		source.pages[0]!.shapes[0]!.children[0]!.name = 'changed';
		expect(target.selection.name).not.toBe('changed');
		expect(Object.isFrozen(target.selection)).toBe(true);
		expect(visioTextOccurrenceSelection(source, { ...occurrence, end: 4 })).toBeUndefined();
		expect(visioTextOccurrenceSelection(source, { ...occurrence, shapeId: '9' })).toBeUndefined();
		expect(visioTextOccurrenceSelection(source, { ...occurrence, pageId: '9' })).toBeUndefined();
	});
	it('uses nonoverlapping UTF16 offsets and literal replacement dollars', () => {
		const source = model('😀aaa A aa');
		const plan = visioTextReplacePlan(source, { ...request, query: 'aa', replacement: '$&' });
		expect(plan.occurrences).toEqual([
			{ pageId: '0', shapeId: '1', start: 2, end: 4 },
			{ pageId: '0', shapeId: '1', start: 8, end: 10 },
		]);
		expect(plan.replacementCount).toBe(2);
		expect(rangeEdit(plan.edits[0]!).ranges).toEqual([
			{ start: 2, end: 4, text: '$&' },
			{ start: 8, end: 10, text: '$&' },
		]);
		expect(source.pages[0]!.shapes[0]!.text.plainText).toBe('😀aaa A aa');
	});
	it('searches beyond old Find preview and per-shape index truncation', () => {
		const source = model('x'.repeat(40000) + 'a');
		expect(visioTextReplaceOccurrences(source, request)[0]!.start).toBe(40000);
		expect(rangeEdit(visioTextReplacePlan(source, request).edits[0]!).ranges).toEqual([
			{ start: 40000, end: 40001, text: 'X' },
		]);
	});
	it('distinguishes page-local IDs and walks current/all pages in source order', () => {
		const source = model('a');
		source.pages.push({ ...source.pages[0]!, id: '2', shapes: [shape('1', 'a')] });
		expect(visioTextReplacePlan(source, request).edits.map((edit) => edit.pageId)).toEqual([
			'0',
			'2',
		]);
		expect(
			visioTextReplacePlan(source, { ...request, scope: 'current-page', pageId: '2' }).edits,
		).toHaveLength(1);
	});
	it('uses ordered selected group subtrees and deduplicates overlapping roots', () => {
		const source = model('a', 'a');
		const group = source.pages[0]!.shapes[0]!;
		group.kind = 'group';
		group.children = [shape('3', 'a'), shape('4', 'a')];
		const plan = visioTextReplacePlan(source, {
			...request,
			scope: 'selection',
			selection: [{ id: '2' }, { id: '3' }, { id: '1' }, { id: '4' }],
		});
		expect(plan.edits.map((edit) => edit.shapeId)).toEqual(['2', '3', '1', '4']);
	});
	it('accepts explicit background-page selection and rejects nonexistent selections', () => {
		const source = model('a');
		source.pages.push({ ...source.pages[0]!, id: '2', shapes: [shape('1', 'a')] });
		expect(
			visioTextReplacePlan(source, {
				...request,
				scope: 'selection',
				selection: [{ id: '1', pageId: '2' }],
			}).edits[0]!.pageId,
		).toBe('2');
		expect(() =>
			visioTextReplacePlan(source, { ...request, scope: 'selection', selection: [{ id: '9' }] }),
		).toThrow();
	});
	it('replaces current then finds the next occurrence after inserted text and wraps', () => {
		const source = model('a a', 'a');
		const matches = visioTextReplaceOccurrences(source, request);
		const plan = visioTextReplacePlan(source, {
			...request,
			mode: 'current',
			replacement: 'aaaa',
			current: matches[0]!,
		});
		expect(rangeEdit(plan.edits[0]!).ranges).toEqual([{ start: 0, end: 1, text: 'aaaa' }]);
		expect(plan.nextOccurrence).toEqual({ pageId: '0', shapeId: '1', start: 5, end: 6 });
		const last = visioTextReplacePlan(source, {
			...request,
			mode: 'current',
			current: matches[2]!,
		});
		expect(last.nextOccurrence).toEqual(matches[0]);
	});
	it('deletes current and handles zero matches without commands', () => {
		const source = model('a a');
		const plan = visioTextReplacePlan(source, { ...request, mode: 'current', replacement: '' });
		expect(rangeEdit(plan.edits[0]!).ranges).toEqual([{ start: 0, end: 1, text: '' }]);
		expect(plan.nextOccurrence!.start).toBe(1);
		expect(visioTextReplacePlan(model('none'), request).edits).toEqual([]);
	});
	it('keeps matched no-op commands for source protection and unsupported-text checks', () => {
		expect(visioTextReplacePlan(model('a'), { ...request, replacement: 'a' }).edits).toHaveLength(
			1,
		);
	});
	it('freezes plans and does not retain caller request or output command objects', () => {
		const compileTimeReadonly = (edit: VisioTextReplacePlanEdit) => {
			if (edit.type === 'replace-text-ranges') {
				// @ts-expect-error Plans own deeply readonly range elements; commands are mutable clones.
				edit.ranges[0]!.text = 'wrong';
			}
		};
		expect(compileTimeReadonly).toBeTypeOf('function');
		const source = model('a'),
			input = { ...request, scope: 'selection' as const, selection: [{ id: '1' }] };
		const plan = visioTextReplacePlan(source, input);
		input.selection[0]!.id = '9';
		input.replacement = 'wrong';
		const edits = visioTextReplaceCommands(source, plan);
		if (edits[0]!.type !== 'replace-text-ranges') throw new Error('Expected range edit');
		edits[0]!.ranges[0]!.text = 'wrong';
		expect(rangeEdit(visioTextReplaceCommands(source, plan)[0]!).ranges[0]!.text).toBe('X');
		expect(Object.isFrozen(rangeEdit(plan.edits[0]!).ranges)).toBe(true);
		expect(Object.isFrozen(rangeEdit(plan.edits[0]!).ranges[0])).toBe(true);
		expect(Object.isFrozen(plan)).toBe(true);
		expect(Object.isFrozen(plan.occurrences[0])).toBe(true);
		expect(Object.isFrozen(plan.edits[0])).toBe(true);
	});
	it('rejects forged plans and exact-scope text/identity/order staleness', () => {
		const source = model('a', 'a'),
			plan = visioTextReplacePlan(source, request);
		expect(() => visioTextReplaceCommands(source, { ...plan })).toThrow(/not produced/);
		source.pages[0]!.shapes.reverse();
		expect(() => visioTextReplaceCommands(source, plan)).toThrow(/stale/);
		source.pages[0]!.shapes.reverse();
		source.pages[0]!.shapes[1]!.text.plainText = 'b';
		expect(() => visioTextReplaceCommands(source, plan)).toThrow(/stale/);
	});
});

describe('replacement input and allocation bounds', () => {
	it('reads a host selection array length once before copying within its cap', () => {
		let reads = 0;
		const selection = new Proxy([{ id: '1' }], {
			get(target, key, receiver) {
				if (key === 'length' && ++reads > 1) throw new Error('length was reread');
				return Reflect.get(target, key, receiver);
			},
		});
		expect(
			visioTextReplacePlan(model('a'), { ...request, scope: 'selection', selection })
				.replacementCount,
		).toBe(1);
		expect(reads).toBe(1);
	});
	it('accepts a leaf at depth64 and refuses deeper source trees', () => {
		const source = model(''),
			root = source.pages[0]!.shapes[0]!;
		let last = root;
		for (let index = 1; index <= 64; index++) {
			const child = shape(String(index + 1), '');
			last.children = [child];
			last = child;
		}
		last.text.plainText = 'a';
		expect(visioTextReplacePlan(source, request).replacementCount).toBe(1);
		last.children = [shape('66', 'a')];
		expect(() => visioTextReplacePlan(source, request)).toThrow(/shape list/);
	});
	it.each([
		{ query: '' },
		{ query: 'x'.repeat(257) },
		{ matchCase: false },
		{ replacement: 'x'.repeat(32769) },
		{ replacement: '\r' },
		{ replacement: '\ud800' },
		{ mode: 'invalid' },
		{ scope: 'invalid' },
		{ current: { pageId: '0', shapeId: '1', start: -1, end: 1 } },
		{ current: { pageId: '0', shapeId: '1', start: 1, end: 2 } },
	])('refuses invalid options %#', (patch) => {
		expect(() =>
			visioTextReplacePlan(model('a'), { ...request, ...patch } as VisioTextReplaceRequest),
		).toThrow();
	});
	it('rejects duplicate page/shape identities and cycles', () => {
		const source = model('a', 'a');
		source.pages[0]!.shapes[1]!.id = '1';
		expect(() => visioTextReplacePlan(source, request)).toThrow();
		const cyclic = model('a');
		cyclic.pages[0]!.shapes[0]!.children = [cyclic.pages[0]!.shapes[0]!];
		expect(() => visioTextReplacePlan(cyclic, request)).toThrow();
		const duplicate = model('a');
		duplicate.pages.push(duplicate.pages[0]!);
		expect(() => visioTextReplacePlan(duplicate, request)).toThrow();
	});
	it('refuses aggregate source, match, command and output overflow before partial results', () => {
		expect(() => visioTextReplacePlan(model('x'.repeat(2_000_000), 'a'), request)).toThrow(
			/aggregate/,
		);
		expect(() => visioTextReplacePlan(model('a'.repeat(10001)), request)).toThrow(/occurrence/);
		expect(() => visioTextReplacePlan(model(...Array<string>(1001).fill('a')), request)).toThrow(
			/command/,
		);
		expect(() =>
			visioTextReplacePlan(model('a'.repeat(40)), { ...request, replacement: 'X'.repeat(32768) }),
		).toThrow(/output/);
	});
	it('bounds expected source and replacement payload even when final output fits', () => {
		const query = 'a'.repeat(256);
		expect(() =>
			visioTextReplacePlan(model(query.repeat(40) + 'x'.repeat(979760)), {
				...request,
				query,
				replacement: 'b'.repeat(256),
			}),
		).toThrow(/input.*aggregate/);
	});
});
