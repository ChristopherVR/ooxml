import { afterEach, describe, expect, it, vi } from 'vitest';
import { Node, Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import type { DecorationSet } from 'prosemirror-view';
import type { TextRun } from '../model';
import { markSpecs } from './schema-marks';
import { runToInlineNodes } from './run-adapter';
import { symbolFontPlugin } from './symbol-font';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*', toDOM: () => ['p', 0] },
		text: { group: 'inline' },
	},
	marks: markSpecs,
});
const symbol = { font: 'Wingdings', char: 'F0FE' };
const plugin = symbolFontPlugin();

function stateOf(paragraphs: TextRun[][]): EditorState {
	const doc = schema.node(
		'doc',
		null,
		paragraphs.map((runs) =>
			schema.node(
				'paragraph',
				null,
				runs.flatMap((run) => runToInlineNodes(run, schema)),
			),
		),
	);
	return EditorState.create({ schema, doc, plugins: [plugin] });
}

/** Each decorated stretch of text, with the font it is drawn in. */
const decorated = (state: EditorState) =>
	(plugin.getState(state) as DecorationSet).find().map((decoration) => ({
		text: state.doc.textBetween(decoration.from, decoration.to),
		style: (decoration as unknown as { type: { attrs: { style: string } } }).type.attrs.style,
	}));

afterEach(() => vi.restoreAllMocks());

describe('symbol font in the editor', () => {
	it('draws each symbol in its font and follows typing, deleting and formatting', () => {
		let state = stateOf([[{ text: 'Done \uF0FE ok', symbol }], [{ text: 'Other' }]]);
		const font = 'font-family:"Wingdings"';
		expect(decorated(state)).toEqual([{ text: '\uF0FE', style: font }]);
		// Typing elsewhere moves the decoration with its character.
		state = state.apply(state.tr.insertText('Not ', 1));
		expect(decorated(state)).toEqual([{ text: '\uF0FE', style: font }]);
		// A symbol typed into the run is drawn in the symbol font too.
		const end = state.doc.firstChild!.nodeSize - 1;
		state = state.apply(state.tr.insertText('\uF0FE', end));
		expect(decorated(state)).toHaveLength(2);
		state = state.apply(state.tr.delete(end, end + 1));
		expect(decorated(state)).toEqual([{ text: '\uF0FE', style: font }]);
		// Removing the run's properties turns the symbol back into ordinary text.
		state = state.apply(state.tr.removeMark(0, state.doc.content.size, schema.marks.runProperties));
		expect(decorated(state)).toEqual([]);
	});

	it('scans only the text a change touched', () => {
		const long = Array.from({ length: 300 }, (_, index) => [{ text: `Paragraph ${index}` }]);
		let state = stateOf([[{ text: '\uF0FE', symbol }], ...long]);
		const scan = vi.spyOn(Node.prototype, 'nodesBetween');
		state = state.apply(state.tr.insertText('x', state.doc.content.size - 2));
		const widest = Math.max(0, ...scan.mock.calls.map(([from, to]) => to - from));
		expect(widest).toBeLessThan(20);
		expect(decorated(state)).toHaveLength(1);
	});
});
