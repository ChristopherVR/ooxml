import { expect, test } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
import { createDocument, loadDocx, saveDocx } from 'ooxml-core/docx';
import { loadWordReferences, WordYjsCollaboration } from 'ooxml-core/docx/ui';
import { docToModel, modelToDoc } from '../model-adapter';
import { schema } from '../schema';
import { bibliographyOf, setBibliography } from './edits';

test('shares bibliography edits and undo with a second peer and saves that peer’s sources', async () => {
	const references = await loadWordReferences();
	const model = createDocument();
	const initial = modelToDoc(model);
	const hub = createMemoryHub();
	const cleanups: (() => void)[] = [];
	const sessions = [true, false].map((create) =>
		createCollabSession({
			roomId: 'references',
			provider: transportProvider({ transport: hub.createTransport('references') }),
			user: { name: create ? 'Author' : 'Reader' },
			heartbeatMs: 0,
			teardown: false,
		}),
	);
	for (const session of sessions) cleanups.unshift(() => session.destroy());
	const peer = (session: (typeof sessions)[number], create: boolean) => {
		const binding = new WordYjsCollaboration(session, initial, {
			documentId: 'source',
			initializeIfEmpty: create,
		});
		cleanups.unshift(() => binding.destroy());
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create(binding.state(schema)),
		});
		cleanups.unshift(() => view.destroy());
		return { binding, view };
	};
	try {
		const a = peer(sessions[0]!, true);
		const b = peer(sessions[1]!, false);
		const tr = a.view.state.tr;
		const bibliography = {
			style: 'MLA',
			sources: [{ tag: 'Doe20', type: 'Book', fields: { Title: 'Field studies' } }],
		} as const;
		setBibliography(references, tr, model, { ...bibliography, sources: [...bibliography.sources] });
		a.view.dispatch(tr);
		const joined = docToModel(b.view.state.doc, model);
		expect(bibliographyOf(references, joined)).toEqual(bibliography);
		const reopened = (await loadDocx(await saveDocx(joined))).model;
		expect(bibliographyOf(references, reopened)).toMatchObject(bibliography);
		expect(a.binding.undo()).toBe(true);
		expect(bibliographyOf(references, docToModel(b.view.state.doc, joined))?.sources).toEqual([]);
	} finally {
		cleanups.forEach((cleanup) => cleanup());
	}
});
