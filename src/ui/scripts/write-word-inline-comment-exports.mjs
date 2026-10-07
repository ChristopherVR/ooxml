// Run with Bun after building core. Native checks reopen these synthetic exports.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { EditorState, TextSelection } from 'prosemirror-state';
import { loadDocx } from 'ooxml-core/docx';
import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
import { addComment, WordYjsCollaboration } from 'ooxml-core/docx/ui';
import { modelToDoc, docToModel } from '../src/docx/model-adapter.ts';

const output = process.argv[2];
if (!output) throw new Error('Provide an output directory');
await mkdir(output, { recursive: true });
for (const name of ['picture', 'note', 'break', 'field', 'line-break']) {
	const loaded = await loadDocx(
		await readFile(
			new URL(
				`../../core/docx/__fixtures__/review-advanced-object-formatting/${name}-before.docx`,
				import.meta.url,
			),
		),
	);
	const host = {
		state: EditorState.create({ doc: modelToDoc(loaded.model) }),
		editable: true,
		dispatch(tr) {
			this.state = this.state.apply(tr);
		},
	};
	let pos = -1;
	host.state.doc.descendants((node, position) => {
		if (pos < 0 && node.isInline && !node.isText) pos = position;
	});
	if (pos < 0) throw new Error(`Missing ${name} element`);
	host.dispatch(host.state.tr.setSelection(TextSelection.create(host.state.doc, pos, pos + 1)));
	const comments = [
		addComment(host, 'Ada', 'A comment', () => 'a'),
		addComment(host, 'Bob', 'B comment', () => 'b'),
	];
	if (comments.some((comment) => !comment)) throw new Error(`Cannot anchor ${name}`);
	const hub = createMemoryHub();
	const sessions = [0, 1].map((index) =>
		createCollabSession({
			roomId: 'exports',
			provider: transportProvider({ transport: hub.createTransport('exports') }),
			user: { name: String(index) },
			heartbeatMs: 0,
			teardown: false,
		}),
	);
	const bindings = sessions.map(
		(session, index) =>
			new WordYjsCollaboration(session, host.state.doc, {
				documentId: 'source',
				initializeIfEmpty: index === 0,
				initialComments: comments,
			}),
	);
	const views = bindings.map((binding) => ({
		state: EditorState.create(binding.state(host.state.schema)),
		editable: true,
	}));
	try {
		const save = async (mode) => {
			const binding = bindings[1];
			const model = docToModel(binding.state(host.state.schema).doc, {
				...loaded.model,
				comments: binding.comments.all(),
			});
			await writeFile(resolve(output, `${name}-${mode}.docx`), await loaded.save(model));
		};
		await save('comments');
		for (const [index, id] of ['a', 'b'].entries()) {
			if (!bindings[index].comments.delete(views[index], id))
				throw new Error('Cannot delete source comment');
		}
		await save('deleted');
		if (!bindings[0].undo()) throw new Error('Missing local deletion undo');
		await save('restored');
	} finally {
		for (const binding of bindings) binding.destroy();
		for (const session of sessions) session.destroy();
	}
}
console.log(resolve(output));
