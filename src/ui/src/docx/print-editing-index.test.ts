// @vitest-environment jsdom
import { expect, it, vi } from 'vitest';
import { Schema } from 'prosemirror-model';
import { PrintEditingIndex } from './print-editing-index';
const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'text*', attrs: { id: { default: null } } },
		text: {},
	},
});
it('reindexes replaced page render roots while preserving the unchanged immutable document index', () => {
	const doc = schema.node('doc', null, [
		schema.node('paragraph', { id: 'body' }, schema.text('Text')),
	]);
	const pages = document.createElement('div');
	pages.innerHTML =
		'<div><div class="dve-print-line" data-block-id="body" data-source-start="0" data-source-end="4"></div></div>';
	const index = new PrintEditingIndex(pages),
		scan = vi.spyOn(doc, 'descendants'),
		query = vi.spyOn(pages, 'querySelectorAll');
	index.update(doc);
	const first = index.block('body')[0]!;
	pages.innerHTML =
		'<div><div class="dve-print-line" data-block-id="footnote"></div><div class="dve-print-line" data-block-id="body" data-source-start="0" data-source-end="2"></div><div class="dve-print-line" data-block-id="body" data-source-start="2" data-source-end="4"></div></div>';
	index.update(doc);
	expect(index.block('body')).toHaveLength(2);
	expect(index.block('body')).not.toContain(first);
	expect(index.block('footnote')).toHaveLength(0);
	expect(index.adjacent(index.block('body')[0]!, 1)).toBe(index.block('body')[1]);
	expect(scan).toHaveBeenCalledTimes(1);
	expect(query).toHaveBeenCalledTimes(2);
	expect([...index.between(3, 3)]).toEqual([{ id: 'body', start: 1, end: 5 }]);
	index.clear();
	expect(index.position('body', 0)).toBeUndefined();
	vi.restoreAllMocks();
});

it('targets physical page starts and nearest screen lines without selecting unrelated stories', () => {
	const doc = schema.node(
		'doc',
		null,
		Array.from({ length: 3 }, (_, i) =>
			schema.node('paragraph', { id: 'p' + i }, schema.text('Text')),
		),
	);
	const pages = document.createElement('div');
	pages.innerHTML = Array.from(
		{ length: 3 },
		(_, i) =>
			`<div class="dve-print-page"><div class="dve-print-line" data-block-id="p${i}"></div><div class="dve-print-line" data-block-id="footnote"></div></div>`,
	).join('');
	[...pages.children].forEach((page, i) => {
		vi.spyOn(page, 'getBoundingClientRect').mockReturnValue(new DOMRect(0, i * 110, 100, 100));
		vi.spyOn(page.firstElementChild!, 'getBoundingClientRect').mockReturnValue(
			new DOMRect(10, i * 110 + 10, 80, 20),
		);
	});
	const index = new PrintEditingIndex(pages);
	index.update(doc);
	expect(index.pageStep(index.block('p0')[0]!, 1)).toBe(index.block('p1')[0]);
	expect(index.pageStep(index.block('p2')[0]!, 1)).toBe(index.block('p2')[0]);
	expect(index.nearest(20, 135)).toBe(index.block('p1')[0]);
	expect(index.nearest(20, 300)).toBe(index.block('p2')[0]);
	index.clear();
	vi.restoreAllMocks();
});
