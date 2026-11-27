// @vitest-environment jsdom
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it } from 'vitest';
import { browserClickCountPlugin } from './click-count';
import { schema } from './schema';

type Input = { lastClick: { type: string } };
const views: EditorView[] = [];

function editor(): EditorView {
	const doc = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('zq body ')])]);
	const view = new EditorView(document.body.appendChild(document.createElement('div')), {
		state: EditorState.create({ doc, plugins: [browserClickCountPlugin()] }),
	});
	// No layout in jsdom: the press is counted, then finds no position and stops there.
	view.posAtCoords = () => null;
	views.push(view);
	return view;
}

const press = (view: EditorView, detail: number) =>
	view.dom.dispatchEvent(
		new MouseEvent('mousedown', {
			bubbles: true,
			cancelable: true,
			clientX: 2,
			clientY: 4,
			detail,
		}),
	);
const lastClick = (view: EditorView) => (view as unknown as { input: Input }).input.lastClick.type;

afterEach(() => {
	for (const view of views.splice(0)) {
		view.dom.parentElement?.remove();
		view.destroy();
	}
});

describe('browserClickCountPlugin', () => {
	it('keeps two separate clicks at one place single clicks, as the browser counts them', () => {
		// Clicking into the body, Find walking the stories, then clicking the same place again
		// within ProseMirror's 500 ms selected the space under the pointer as a double click would.
		const view = editor();
		press(view, 1);
		press(view, 1);
		expect(lastClick(view)).toBe('singleClick');
	});

	it('still takes the browser double and triple clicks', () => {
		const view = editor();
		press(view, 1);
		press(view, 2);
		expect(lastClick(view)).toBe('doubleClick');
		press(view, 3);
		expect(lastClick(view)).toBe('tripleClick');
	});
});
