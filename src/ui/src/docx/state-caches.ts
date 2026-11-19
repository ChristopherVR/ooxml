import { history } from 'prosemirror-history';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';

let blank: { state: EditorState; view: EditorView } | undefined;

/**
 * prosemirror-history (the plugins of the last state it recorded a change in) and
 * prosemirror-view (the last state it measured a textblock end in, for arrow keys and Backspace)
 * each keep one editor state in a module variable. Through its plugins that state holds the editor
 * that made them, so a removed editor would live until another one is edited. Recording and
 * measuring a small state of our own moves both caches off it.
 */
export function releaseStateCaches(): void {
	if (!blank) {
		const schema = new Schema({ nodes: { doc: { content: 'text*' }, text: {} } });
		const state = EditorState.create({ schema, plugins: [history()] });
		blank = { state, view: new EditorView(null, { state, editable: () => false }) };
	}
	blank.state.apply(blank.state.tr.insertText(' '));
	blank.view.endOfTextblock('forward');
}
