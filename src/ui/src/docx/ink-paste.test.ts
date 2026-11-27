// @vitest-environment jsdom
import { DOMSerializer } from 'prosemirror-model';
import { EditorState, NodeSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { inkPastePlugin } from './ink-paste';
import { schema } from './schema';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

const ink = () =>
	schema.nodes.image!.create({
		unsupported: 'Ink',
		widthPx: 40,
		heightPx: 20,
		anchored: true,
		placement: JSON.stringify({
			wrap: 'none',
			relativeFrom: 'page',
			relativeFromV: 'page',
			offsetXPx: 100,
			offsetYPx: 50,
		}),
		ink: JSON.stringify({
			strokes: [{ path: 'M 0 0 L 30 10', color: '#000000', width: 2, opacity: 1 }],
			width: 40,
			height: 20,
		}),
	});

describe('pasted ink', () => {
	it('is selected and moved off a drawing it would cover', () => {
		// jsdom has no ClipboardEvent, which pasteHTML creates.
		vi.stubGlobal(
			'ClipboardEvent',
			class extends Event {
				clipboardData = null;
			},
		);
		const doc = schema.node('doc', null, [
			schema.nodes.paragraph!.create({ id: 'p0' }, [ink(), schema.text('text')]),
		]);
		const host = document.createElement('div');
		document.body.append(host);
		const view = new EditorView(host, {
			state: EditorState.create({ doc, schema, plugins: [inkPastePlugin()] }),
		});
		const holder = document.createElement('div');
		holder.append(DOMSerializer.fromSchema(schema).serializeNode(ink()));
		view.dispatch(
			view.state.tr.setSelection(
				NodeSelection.near(view.state.doc.resolve(view.state.doc.content.size - 1)),
			),
		);
		view.pasteHTML(holder.innerHTML);
		const selection = view.state.selection;
		expect(selection).toBeInstanceOf(NodeSelection);
		const pasted = (selection as NodeSelection).node;
		expect(pasted.attrs.ink).toBeTruthy();
		expect(JSON.parse(pasted.attrs.placement)).toMatchObject({ offsetXPx: 112, offsetYPx: 62 });
		let count = 0;
		view.state.doc.descendants((node) => {
			if (node.attrs.ink) count++;
		});
		expect(count).toBe(2);
		view.destroy();
	});
});
