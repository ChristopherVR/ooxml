// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView, DecorationSet } from 'prosemirror-view';
import { describe, expect, it, vi } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import {
	contentControlPlugin,
	controlSpans,
	typeOverPlaceholder,
	updateControlProperties,
} from 'ooxml-core/docx/ui';
import { contentControlTagWidgets } from './content-control-tag-widgets';
import { contentControlViewPlugin } from './content-control-view';
import { docToModel, modelToDoc } from './model-adapter';

async function load(name: string) {
	return loadDocx(
		await readFile(
			new URL(
				'../../../core/docx/layout/fixtures/word-layout-reports/' + name + '.docx',
				import.meta.url,
			),
		),
	);
}

async function withView(name: string, run: (view: EditorView) => Promise<void> | void) {
	const loaded = await load(name);
	const host = document.body.appendChild(document.createElement('div'));
	const view = new EditorView(host, {
		state: EditorState.create({
			doc: modelToDoc(loaded.model),
			plugins: [contentControlPlugin(), contentControlViewPlugin()],
		}),
	});
	try {
		await run(view);
	} finally {
		view.destroy();
		host.remove();
	}
}

function select(view: EditorView, pos: number) {
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
}

function anchorIsOnContentSide(view: EditorView, tag: HTMLElement) {
	const { node, offset } = view.domAtPos(view.state.selection.from);
	const anchor = document.createRange();
	anchor.setStart(node, offset);
	anchor.collapse(true);
	const boundary = document.createRange();
	boundary.selectNode(tag);
	const start = tag.dataset.tagEdge === 'start';
	boundary.collapse(!start);
	const order = anchor.compareBoundaryPoints(Range.START_TO_START, boundary);
	return start ? order >= 0 : order <= 0;
}

function inlineSpan(view: EditorView) {
	const span = controlSpans(view.state.doc).find((entry) => entry.level === 'inline');
	if (!span) throw new Error('Missing inline control');
	return span;
}

describe('content control tag caret', () => {
	it('brackets the DOM anchor at paragraph, row and cell edges', async () => {
		await withView('block-content-control-tags', (view) => {
			const widgets = contentControlTagWidgets(view.state.doc).find();
			expect(widgets).toHaveLength(6);
			const sides = (edge: string) =>
				widgets.filter(({ spec }) => spec.edge === edge).map(({ spec }) => Math.sign(spec.side));
			expect(sides('start')).toEqual([-1, -1, -1]);
			expect(sides('end')).toEqual([1, 1, 1]);
			const tags = view.dom.querySelectorAll<HTMLElement>('.dve-sdt-tag');
			expect(tags).toHaveLength(6);
			expect(
				[...tags].filter((tag) => tag.dataset.tagEdge === 'end').map((tag) => tag.dataset.tagTitle),
			).toEqual(['Block', 'Cells', 'Rows']);
			for (const tag of tags) {
				select(view, view.posAtDOM(tag, 0));
				expect(anchorIsOnContentSide(view, tag)).toBe(true);
			}
		});
	});

	it('keeps the caret on its approach side of an inline tag', async () => {
		await withView('issue-349-content-control-tags', (view) => {
			const { from, to } = inlineSpan(view);
			const [start, end] = view.dom.querySelectorAll<HTMLElement>('.dve-sdt-tag');
			if (!start || !end) throw new Error('Missing inline tags');
			const tagAt = (pos: number, edge: string) =>
				[...view.dom.querySelectorAll<HTMLElement>('.dve-sdt-tag')].find(
					(tag) => tag.dataset.tagEdge === edge && view.posAtDOM(tag, 0) === pos,
				);
			for (const [prior, pos, edge, inside] of [
				[from - 1, from, 'start', false],
				[from + 1, from, 'start', true],
				[to + 1, to, 'end', false],
				[to - 1, to, 'end', true],
			] as const) {
				select(view, prior);
				select(view, pos);
				const tag = tagAt(pos, edge);
				if (!tag) throw new Error('Missing boundary tag');
				expect(anchorIsOnContentSide(view, tag)).toBe(inside);
				const decorations = view.someProp('decorations', (draw) => draw(view.state));
				if (!(decorations instanceof DecorationSet)) throw new Error('Missing widgets');
				const [widget] = decorations.find(pos, pos, (spec) => spec.inline && spec.edge === edge);
				expect(Math.sign(widget?.spec.side)).toBe((edge === 'start') === inside ? -1 : 1);
				const marks = view.state.storedMarks ?? view.state.selection.$from.marks();
				expect(marks.some((mark) => mark.type.name === 'contentControl')).toBe(inside);
			}
		});
	});

	it('clicks through to the content side and saves typed text inside the inline control', async () => {
		const loaded = await load('issue-349-content-control-tags');
		await withView('issue-349-content-control-tags', async (view) => {
			const span = inlineSpan(view);
			const tag = view.dom.querySelector<HTMLElement>('.dve-sdt-tag');
			if (!tag) throw new Error('Missing start tag');
			view.posAtCoords = () => null;
			tag.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
			expect(view.state.selection.from).toBe(span.from);
			view.dispatch(view.state.tr.insertText('Lorem '));
			const endTag = view.dom.querySelector<HTMLElement>('.dve-sdt-tag[data-tag-edge=end]');
			if (!endTag) throw new Error('Missing end tag');
			endTag.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
			expect(view.state.selection.from).toBe(span.to + 'Lorem '.length);
			view.dispatch(view.state.tr.insertText(' ipsum'));
			const saved = await loaded.save(docToModel(view.state.doc, loaded.model));
			const reopened = modelToDoc((await loadDocx(saved)).model);
			const edited = controlSpans(reopened).find(({ control }) => control.id === span.control.id);
			expect(edited?.control.appearance).toBe('tags');
			expect(edited && reopened.textBetween(edited.from, edited.to)).toBe('Lorem alpha ipsum');
		});
	});

	it('keeps outside typing beside a tagged placeholder and protects a locked content edge', async () => {
		await withView('issue-349-content-control-tags', (view) => {
			const span = inlineSpan(view);
			view.dispatch(
				updateControlProperties(view.state, span, { ...span.control, showingPlaceholder: true }),
			);
			select(view, span.from - 1);
			select(view, span.from);
			expect(Boolean(typeOverPlaceholder(view.state, span.from, 'Lorem'))).toBe(false);
			view.dispatch(view.state.tr.insertText('Lorem '));
			const moved = inlineSpan(view);
			expect(view.state.doc.textBetween(moved.from, moved.to)).toBe('alpha');
			view.dispatch(
				updateControlProperties(view.state, moved, {
					...moved.control,
					lock: 'contentLocked',
					showingPlaceholder: false,
				}),
			);
			select(view, moved.to - 1);
			select(view, moved.to);
			const before = view.state.doc;
			view.dispatch(view.state.tr.insertText('Lorem'));
			expect(view.state.doc).toBe(before);
		});
	});
});

describe('content control tag updates', () => {
	it.each([
		['issue-349-content-control-tags', 240],
		['issue-349-content-control-tags-control', 0],
	] as const)('maps ordinary text edits in %s without a document walk', async (name, count) => {
		const doc = modelToDoc((await load(name)).model);
		const plugin = contentControlViewPlugin();
		let state = EditorState.create({ doc, plugins: [plugin] });
		const widgets = (current: EditorState) => {
			const decorations = plugin.props.decorations?.call(plugin, current);
			if (!(decorations instanceof DecorationSet)) throw new Error('Missing decoration set');
			return decorations.find().filter(({ spec }) => spec.side !== undefined);
		};
		const original = widgets(state)
			.map(({ from }) => from)
			.sort((a, b) => a - b);
		expect(original).toHaveLength(count);
		const span = controlSpans(doc)[0];
		if (!span) throw new Error('Missing test control');
		for (const position of [2, span.from + 2]) {
			const tr = state.tr.insertText('x', position);
			const walk = vi.spyOn(tr.doc, 'descendants');
			try {
				state = state.apply(tr);
				expect(walk).not.toHaveBeenCalled();
			} finally {
				walk.mockRestore();
			}
		}
		expect(
			widgets(state)
				.map(({ from }) => from)
				.sort((a, b) => a - b),
		).toEqual(original.map((from) => from + (from <= span.from ? 1 : 2)));
	});

	it('rebuilds and round trips an appearance edit when the tag set was empty', async () => {
		const loaded = await load('issue-349-content-control-tags-control');
		await withView('issue-349-content-control-tags-control', async (view) => {
			const span = inlineSpan(view);
			expect(view.dom.querySelectorAll('.dve-sdt-tag')).toHaveLength(0);
			const tr = updateControlProperties(view.state, span, { ...span.control, appearance: 'tags' });
			const walk = vi.spyOn(tr.doc, 'descendants');
			try {
				view.dispatch(tr);
				expect(walk).toHaveBeenCalled();
			} finally {
				walk.mockRestore();
			}
			expect(view.dom.querySelectorAll('.dve-sdt-tag')).toHaveLength(2);
			const saved = await loaded.save(docToModel(view.state.doc, loaded.model));
			const reopened = modelToDoc((await loadDocx(saved)).model);
			expect(contentControlTagWidgets(reopened).find()).toHaveLength(2);
		});
	});
});

it('rebuilds when a plain paragraph splits or rejoins a tagged block', async () => {
	await withView('block-content-control-tags', (view) => {
		const pos = view.state.doc.child(0).nodeSize;
		const paragraph = view.state.schema.node('paragraph', null, view.state.schema.text('Tempor'));
		const tags = () => view.dom.querySelectorAll('.dve-sdt-tag');
		expect(tags()).toHaveLength(6);
		view.dispatch(view.state.tr.insert(pos, paragraph));
		expect(tags()).toHaveLength(8);
		view.dispatch(view.state.tr.delete(pos, pos + paragraph.nodeSize));
		expect(tags()).toHaveLength(6);
	});
});

it('responds to a row attribute edit and replacement of a tagged range', async () => {
	await withView('block-content-control-tags', (view) => {
		const rowTag = view.dom.querySelector<HTMLElement>('[data-tag-title=Rows]');
		if (!rowTag) throw new Error('Missing row tag');
		const $tag = view.state.doc.resolve(view.posAtDOM(rowTag, 0));
		let depth = $tag.depth;
		while (depth > 0 && $tag.node(depth).type.name !== 'tableRow') depth--;
		const table = $tag.node(depth - 1);
		const tr = view.state.tr;
		const tableStart = $tag.start(depth - 1);
		table.forEach((row, offset) => {
			tr.setNodeAttribute(tableStart + offset, 'properties', {
				...row.attrs.properties,
				contentControl: { ...row.attrs.properties.contentControl, appearance: 'hidden' },
			});
		});
		view.dispatch(tr);
		expect(view.dom.querySelector('[data-tag-title=Rows]')).toBeNull();
		expect(view.dom.querySelector('[data-tag-title=Block]')).not.toBeNull();
		view.dispatch(view.state.tr.delete(0, view.state.doc.child(0).nodeSize));
		expect(view.dom.querySelector('[data-tag-title=Block]')).not.toBeNull();
		expect(view.dom.querySelector('p')?.textContent).toBe('dolor sit');
	});
});

it('shows a pasted tagged control and drops both widgets when its marked text is deleted', async () => {
	const source = modelToDoc((await load('issue-349-content-control-tags')).model);
	const span = controlSpans(source)[0];
	if (!span) throw new Error('Missing tag to paste');
	await withView('issue-349-content-control-tags-control', (view) => {
		const pos = 3;
		const slice = source.slice(span.from, span.to);
		view.dispatch(view.state.tr.replace(pos, pos, slice));
		expect(view.dom.querySelectorAll('.dve-sdt-tag')).toHaveLength(2);
		view.dispatch(view.state.tr.delete(pos, pos + slice.size));
		expect(view.dom.querySelectorAll('.dve-sdt-tag')).toHaveLength(0);
	});
});

it('splits and rejoins tags at an unmarked inline atom between marked runs', async () => {
	await withView('issue-349-content-control-tags', (view) => {
		const { from } = inlineSpan(view);
		const { schema } = view.state;
		const bold = schema.marks.bold;
		if (!bold) throw new Error('Missing bold mark');
		view.dispatch(view.state.tr.addMark(from, from + 1, bold.create()));
		const tags = () => view.dom.querySelectorAll('.dve-sdt-tag');
		expect(tags()).toHaveLength(240);
		view.dispatch(view.state.tr.insert(from + 1, schema.node('hardBreak')));
		expect(tags()).toHaveLength(242);
		view.dispatch(view.state.tr.delete(from + 1, from + 2));
		expect(tags()).toHaveLength(240);
	});
});

it('keeps every character typed at the inside end of an inline control inside it', async () => {
	await withView('issue-349-content-control-tags', (view) => {
		const span = inlineSpan(view);
		select(view, span.to - 1);
		select(view, span.to);
		for (const char of 'xyz') view.dispatch(view.state.tr.insertText(char));
		const typed = controlSpans(view.state.doc).find(
			({ control }) => control.id === span.control.id,
		);
		expect(typed && view.state.doc.textBetween(typed.from, typed.to)).toBe('alphaxyz');
	});
});

it('types beside a tagged check box at the end of a paragraph', async () => {
	await withView('issue-349-content-control-tags', (view) => {
		const { schema } = view.state;
		const box = {
			id: '999',
			level: 'inline',
			kind: 'checkbox',
			checked: false,
			appearance: 'tags',
		};
		const mark = schema.marks.contentControl!.create({ data: JSON.stringify(box) });
		view.dispatch(
			view.state.tr.insert(view.state.doc.child(0).nodeSize - 1, schema.text('☐', [mark])),
		);
		const end = view.state.doc.child(0).nodeSize - 1;
		select(view, 2);
		select(view, end);
		expect(typeOverPlaceholder(view.state, end, 'x')).not.toBe('blocked');
		const marks = view.state.storedMarks ?? view.state.selection.$from.marks();
		expect(marks.some((entry) => entry.type.name === 'contentControl')).toBe(false);
	});
});
