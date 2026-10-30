// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { history, redo, undo } from 'prosemirror-history';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { SearchController, findTextMatches } from './search-controller';
import { createSearchPanel } from './search-panel';

const bold = schema.marks.bold.create();
const italic = schema.marks.italic.create();
const paragraph = (id: string, ...children: ReturnType<typeof schema.text>[]) =>
	schema.nodes.paragraph.create({ id }, children);

function makeView(doc: ReturnType<typeof schema.node>) {
	const host = document.createElement('div');
	document.body.append(host);
	let view: EditorView;
	const state = EditorState.create({ doc, plugins: [history()] });
	view = new EditorView(host, {
		state,
		dispatchTransaction(transaction) {
			view.updateState(view.state.apply(transaction));
		},
	});
	return view;
}

afterEach(() => document.body.replaceChildren());

describe('SearchController', () => {
	it('finds across adjacent differently marked runs and replaces without disturbing surrounding marks or paragraph ids', () => {
		const doc = schema.node('doc', null, [
			paragraph(
				'stable-id',
				schema.text('Hel', [bold]),
				schema.text('lo ', [italic]),
				schema.text('HELLO'),
			),
		]);
		const view = makeView(doc);
		const search = new SearchController(() => view);
		expect(search.search('hello').count).toBe(2);
		const first = search.findNext()!;
		const second = search.findNext()!;
		expect(second.from).toBeGreaterThan(first.from);
		expect(search.findPrevious()).toEqual(first);
		expect(search.replaceOne('$&😀')).toBe(true);
		expect(view.state.doc.textContent).toBe('$&😀 HELLO');
		expect(view.state.doc.firstChild?.attrs.id).toBe('stable-id');
		const replacement = view.state.doc.firstChild!.firstChild!;
		expect(replacement.text).toBe('$&😀');
		expect(replacement.marks.map((mark) => mark.type.name)).toContain('bold');
		const remaining = view.state.doc.firstChild!.child(1);
		expect(remaining.text).toBe(' ');
		expect(remaining.marks.map((mark) => mark.type.name)).toContain('italic');
		view.destroy();
	});

	it('uses Unicode case folding without matching half of a surrogate pair or a combining grapheme', () => {
		const text = 'A😀e\u0301 CAFÉ';
		const doc = schema.node('doc', null, paragraph('unicode', schema.text(text)));
		expect(findTextMatches(doc, '😀')).toHaveLength(1);
		expect(findTextMatches(doc, '\ud83d')).toHaveLength(0);
		expect(findTextMatches(doc, 'e')).toHaveLength(0);
		expect(findTextMatches(doc, 'e\u0301')).toHaveLength(1);
		expect(findTextMatches(doc, 'café')).toHaveLength(1);
		expect(findTextMatches(doc, 'café', { matchCase: true })).toHaveLength(0);
	});

	it('finds words that hold soft hyphens, which show only where a line breaks, and replaces them whole', () => {
		const shy = '\u00AD';
		const doc = schema.node(
			'doc',
			null,
			paragraph('soft', schema.text(`Hyphen${shy}ation and hyphen${shy} and ${shy}hyphen`)),
		);
		expect(findTextMatches(doc, 'hyphenation')).toEqual([{ from: 1, to: 13 }]);
		// A soft hyphen at either end of a match is not part of it.
		expect(findTextMatches(doc, 'hyphen').map(({ from, to }) => to - from)).toEqual([6, 6, 6]);
		expect(findTextMatches(doc, `hyphen${shy}ation`)).toHaveLength(1);
		expect(findTextMatches(doc, shy)).toHaveLength(0);
		const view = makeView(doc);
		const search = new SearchController(() => view);
		search.search('hyphenation');
		expect(search.replaceAll('splitting')).toBe(1);
		expect(view.state.doc.textContent).toBe(`splitting and hyphen${shy} and ${shy}hyphen`);
		view.destroy();
	});

	it('does not match across paragraphs, cells, or hard breaks', () => {
		const lineOne = paragraph('one', schema.text('foo'), schema.text('bar', [bold]));
		const lineTwo = paragraph(
			'two',
			schema.text('fo'),
			schema.nodes.hardBreak.create(),
			schema.text('obar'),
		);
		const table = schema.nodes.table.create(null, [
			schema.nodes.tableRow.create(null, [
				schema.nodes.tableCell.create(null, paragraph('left-cell', schema.text('left'))),
				schema.nodes.tableCell.create(null, paragraph('right-cell', schema.text('right'))),
			]),
		]);
		const doc = schema.node('doc', null, [lineOne, lineTwo, table]);
		expect(findTextMatches(doc, 'foobar')).toHaveLength(1);
		expect(findTextMatches(doc, 'fo\nobar')).toHaveLength(0);
		expect(findTextMatches(doc, 'leftright')).toHaveLength(0);
		expect(findTextMatches(doc, 'left')).toHaveLength(1);
	});

	it('replaces all matches in one undoable transaction and emits hard breaks for newline replacement', () => {
		const original = schema.node('doc', null, [
			paragraph('first', schema.text('cat ', [bold]), schema.text('cat', [italic])),
			paragraph('second', schema.text('cat')),
		]);
		const view = makeView(original);
		const search = new SearchController(() => view);
		search.search('cat', { matchCase: true });
		expect(search.replaceAll('dog\ncat')).toBe(3);
		expect(view.state.doc.child(0).textContent).toBe('dog\ncat dog\ncat');
		expect(view.state.doc.child(1).textContent).toBe('dog\ncat');
		const breaks: number[] = [];
		view.state.doc.descendants((node, pos) => {
			if (node.type.name === 'hardBreak') breaks.push(pos);
		});
		expect(breaks).toHaveLength(3);
		const replacedNodes: Array<{ text?: string; marks: string[] }> = [];
		view.state.doc.descendants((node) => {
			if (node.isText && node.text === 'dog')
				replacedNodes.push({ text: node.text, marks: node.marks.map((mark) => mark.type.name) });
		});
		expect(replacedNodes.map((node) => node.marks)).toEqual([['bold'], ['italic'], []]);
		expect(view.state.doc.child(0).attrs.id).toBe('first');
		expect(view.state.doc.child(1).attrs.id).toBe('second');
		undo(view.state, (transaction) => view.dispatch(transaction), view);
		expect(view.state.doc.eq(original)).toBe(true);
		redo(view.state, (transaction) => view.dispatch(transaction), view);
		expect(view.state.doc.child(0).textContent).toBe('dog\ncat dog\ncat');
		expect(view.state.doc.child(1).textContent).toBe('dog\ncat');
		view.setProps({ editable: () => false });
		const unchanged = view.state.doc;
		expect(search.replaceAll('blocked')).toBe(0);
		expect(view.state.doc.eq(unchanged)).toBe(true);
		view.destroy();
	});

	it('treats empty queries safely and supports delete replacement as a literal operation', () => {
		const view = makeView(schema.node('doc', null, paragraph('empty', schema.text('a$a a$a'))));
		const search = new SearchController(() => view);
		expect(search.search('').count).toBe(0);
		expect(search.replaceAll('x')).toBe(0);
		search.search('$a', { matchCase: true });
		expect(search.replaceOne('')).toBe(true);
		expect(view.state.doc.textContent).toBe('a a$a');
		view.destroy();
	});
});

describe('accessible search panel', () => {
	it('exposes search controls, updates its count, disables replacement in read-only mode, and closes on Escape', () => {
		const view = makeView(schema.node('doc', null, paragraph('panel', schema.text('one one'))));
		const onClose = vi.fn();
		const panel = createSearchPanel({ getView: () => view, onClose });
		document.body.append(panel.element);
		panel.open();
		const query = panel.element.querySelector<HTMLInputElement>('[aria-label="Find text"]')!;
		query.value = 'one';
		query.dispatchEvent(new Event('input', { bubbles: true }));
		query.focus();
		query.setSelectionRange(1, 2);
		panel.refresh();
		expect(query.selectionStart).toBe(1);
		expect(query.selectionEnd).toBe(2);
		expect(panel.element.querySelector('[role="status"]')?.textContent).toBe('2 matches');
		for (const label of ['Find previous', 'Find next', 'Replace', 'Replace all', 'Close search'])
			expect(panel.element.querySelector(`[aria-label="${label}"]`)).not.toBeNull();
		expect(panel.element.querySelector('[aria-label="Match case"]')).not.toBeNull();
		view.setProps({ editable: () => false });
		panel.refresh();
		expect(
			panel.element.querySelector<HTMLButtonElement>('[aria-label="Replace all"]')?.disabled,
		).toBe(true);
		query.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(panel.isOpen).toBe(false);
		expect(onClose).toHaveBeenCalledOnce();
		view.destroy();
	});
});

describe('Find and Replace options', () => {
	const doc = (...paragraphs: ReturnType<typeof paragraph>[]) =>
		schema.node('doc', null, paragraphs);

	it('finds across paragraph marks with ^p and joins paragraphs when replacing', () => {
		const view = makeView(
			doc(paragraph('a', schema.text('Hello')), paragraph('b', schema.text('world'))),
		);
		const search = new SearchController(() => view);
		expect(search.search('o^pw').count).toBe(1);
		expect(search.replaceAll(' ')).toBe(1);
		expect(view.state.doc.childCount).toBe(1);
		expect(view.state.doc.textContent).toBe('Hell orld');
		expect(view.state.doc.firstChild?.attrs.id).toBe('a');
		view.destroy();
	});

	it('splits a paragraph where the replacement holds ^p and gives the new paragraph its own id', () => {
		const view = makeView(doc(paragraph('split', schema.text('one two'))));
		const search = new SearchController(() => view);
		search.search(' ');
		expect(search.replaceAll('^p')).toBe(1);
		expect(view.state.doc.childCount).toBe(2);
		expect(view.state.doc.child(0).textContent).toBe('one');
		expect(view.state.doc.child(1).textContent).toBe('two');
		expect(view.state.doc.child(0).attrs.id).toBe('split');
		expect(view.state.doc.child(1).attrs.id).not.toBe('split');
		view.destroy();
	});

	it('replaces with wildcard groups \\1 and \\2 and with ^& and ^t', () => {
		const view = makeView(doc(paragraph('names', schema.text('John Smith'))));
		const search = new SearchController(() => view);
		search.search('(J*) (S*h)', { wildcards: true });
		expect(search.replaceAll('\\2, \\1')).toBe(1);
		expect(view.state.doc.textContent).toBe('Smith, John');
		search.search('Smith', { matchCase: true });
		expect(search.replaceAll('[^&]^t')).toBe(1);
		expect(view.state.doc.textContent).toBe('[Smith]\t, John');
		view.destroy();
	});

	it('finds whole words only and the Special codes for tabs and any digit', () => {
		const view = makeView(doc(paragraph('words', schema.text('the there\tthe 7 x'))));
		const search = new SearchController(() => view);
		expect(search.search('the', { wholeWords: true }).count).toBe(2);
		expect(search.search('the', {}).count).toBe(3);
		expect(search.search('^t', {}).count).toBe(1);
		expect(search.search('^#', {}).count).toBe(1);
		view.destroy();
	});

	it('reports an invalid wildcard pattern and replaces nothing', () => {
		const view = makeView(doc(paragraph('bad', schema.text('a(b'))));
		const search = new SearchController(() => view);
		const status = search.search('(a', { wildcards: true });
		expect(status).toMatchObject({ count: 0, invalid: true });
		expect(search.replaceAll('x')).toBe(0);
		expect(view.state.doc.textContent).toBe('a(b');
		expect(search.search('(a', {}).invalid).toBe(false);
		view.destroy();
	});

	it('does not match across table cells, even with wildcards or ^p', () => {
		const table = schema.nodes.table.create(null, [
			schema.nodes.tableRow.create(null, [
				schema.nodes.tableCell.create(null, paragraph('c1', schema.text('x'))),
				schema.nodes.tableCell.create(null, paragraph('c2', schema.text('y'))),
			]),
		]);
		const cells = schema.node('doc', null, [table]);
		expect(findTextMatches(cells, 'x*y', { wildcards: true })).toHaveLength(0);
		expect(findTextMatches(cells, 'x^py')).toHaveLength(0);
	});

	it('shows the More >> options, disables the wildcard-off options, and inserts Special codes', () => {
		const view = makeView(doc(paragraph('panel', schema.text('tab\there'))));
		const panel = createSearchPanel({ getView: () => view });
		document.body.append(panel.element);
		panel.open();
		const more = panel.element.querySelector<HTMLButtonElement>('[aria-label="More >>"]')!;
		const options = panel.element.querySelector<HTMLElement>('.dve-search-options')!;
		expect(options.hidden).toBe(true);
		more.click();
		expect(options.hidden).toBe(false);
		expect(more.getAttribute('aria-expanded')).toBe('true');
		const wildcards = panel.element.querySelector<HTMLInputElement>(
			'[aria-label="Use wildcards"]',
		)!;
		const whole = panel.element.querySelector<HTMLInputElement>(
			'[aria-label="Find whole words only"]',
		)!;
		wildcards.checked = true;
		wildcards.dispatchEvent(new Event('change'));
		expect(whole.disabled).toBe(true);
		wildcards.checked = false;
		wildcards.dispatchEvent(new Event('change'));
		expect(whole.disabled).toBe(false);
		const query = panel.element.querySelector<HTMLInputElement>('[aria-label="Find text"]')!;
		query.value = 'tab';
		query.dispatchEvent(new Event('input'));
		const special = panel.element.querySelector<HTMLSelectElement>('select')!;
		special.value = '^t';
		special.dispatchEvent(new Event('change'));
		expect(query.value).toBe('tab^t');
		expect(panel.element.querySelector('[role="status"]')?.textContent).toBe('1 match');
		view.destroy();
	});
});
