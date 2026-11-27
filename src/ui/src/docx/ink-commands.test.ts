// @vitest-environment jsdom
import { history, undo } from 'prosemirror-history';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { loadTools } from '../ink/tools';
import { InkLasso } from './draw-lasso';
import { applyInkEdits, fitInk, insertInk } from './ink-commands';
import { inkOf } from './ink-node-view';
import { editsOwnText, historyKey } from './ribbon-history-keys';
import { schema } from './schema';

beforeAll(async () => {
	await loadTools();
});
afterEach(() => document.body.replaceChildren());

function editor() {
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create({ id: 'p0' }, schema.text('text')),
	]);
	const host = document.createElement('div');
	document.body.append(host);
	return new EditorView(host, {
		state: EditorState.create({ doc, schema, plugins: [history()] }),
	});
}

const stroke = (x1: number, y1: number, x2: number, y2: number) => ({
	path: `M ${x1} ${y1} L ${x2} ${y2}`,
	color: '#000000',
	width: 2,
	opacity: 1,
});

/** Ink of two strokes on the page: one at (100, 100), one at (300, 200). */
function drawn(view: EditorView) {
	const pos = insertInk(view, [stroke(100, 100, 140, 120), stroke(300, 200, 340, 240)], {
		blockId: 'p0',
	})!;
	return { pos, node: () => view.state.doc.nodeAt(pos)! };
}

const placement = (node: ReturnType<ReturnType<typeof drawn>['node']>) =>
	JSON.parse(String(node.attrs.placement)) as { offsetXPx: number; offsetYPx: number };

describe('Word ink edits', () => {
	it('fits the box and offsets to the strokes an edit leaves', () => {
		const view = editor();
		const { pos, node } = drawn(view);
		expect(placement(node())).toMatchObject({ offsetXPx: 99, offsetYPx: 99 });
		const ink = inkOf(node())!;
		applyInkEdits(view, [{ pos, strokes: ink.strokes.slice(1) }]);
		// Only the second stroke is left: the box moves to it and shrinks around it.
		expect(placement(node())).toMatchObject({ offsetXPx: 299, offsetYPx: 199 });
		expect(Number(node().attrs.widthPx)).toBeLessThan(50);
		const left = inkOf(node())!;
		expect(left.strokes[0]?.path.startsWith('M 1 1')).toBe(true);
	});

	it('makes an eraser gesture one undo step, its box fitted at the end', () => {
		const view = editor();
		const { pos, node } = drawn(view);
		const before = String(node().attrs.ink);
		const ink = inkOf(node())!;
		const tr = applyInkEdits(view, [{ pos, strokes: ink.strokes.slice(1) }], { fit: false })!;
		// While the gesture runs the box stays as the pages show it.
		expect(placement(node())).toMatchObject({ offsetXPx: 99 });
		// The eraser maps the drawings it touched through each edit (setNodeMarkup replaces the node).
		fitInk(view, [tr.mapping.map(pos, -1)], tr.time);
		expect(placement(node())).toMatchObject({ offsetXPx: 299 });
		undo(view.state, view.dispatch);
		expect(String(node().attrs.ink)).toBe(before);
		expect(placement(node())).toMatchObject({ offsetXPx: 99 });
	});

	it('drops a lasso selection when the document changes', () => {
		const view = editor();
		const lasso = new InkLasso();
		const { pos, node } = drawn(view);
		const sheet = document.createElement('div');
		const element = document.createElement('div');
		sheet.append(element);
		element.getBoundingClientRect = () => new DOMRect(0, 0, Number(node().attrs.widthPx), 200);
		const ink = inkOf(node())!;
		lasso.select(
			view,
			[{ pos, ink, attr: String(node().attrs.ink), element, sheet, blockId: 'p0', indexes: [] }],
			[
				{ x: -5, y: -5 },
				{ x: 500, y: -5 },
				{ x: 500, y: 500 },
				{ x: -5, y: 500 },
			],
		);
		expect(lasso.size).toBe(1);
		expect(lasso.targets(view)).toHaveLength(1);
		view.dispatch(view.state.tr.insertText('x', 1));
		expect(lasso.follow(view.state.doc)).toBe(true);
		expect(lasso.size).toBe(0);
	});
});

describe('Undo from the ribbon', () => {
	it('reads Undo and Redo keys and leaves text boxes their own', () => {
		const key = (init: KeyboardEventInit) => historyKey(new KeyboardEvent('keydown', init));
		expect(key({ key: 'z', ctrlKey: true })).toBe('undo');
		expect(key({ key: 'Z', ctrlKey: true, shiftKey: true })).toBe('redo');
		expect(key({ key: 'y', metaKey: true })).toBe('redo');
		expect(key({ key: 'z' })).toBeUndefined();
		const input = document.createElement('input');
		document.body.append(input);
		let owns: boolean | undefined;
		input.addEventListener('keydown', (event) => (owns = editsOwnText(event)));
		input.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true }));
		expect(owns).toBe(true);
	});
});
