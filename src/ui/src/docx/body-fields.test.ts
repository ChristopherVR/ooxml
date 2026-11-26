import { createDocument } from 'ooxml-core/docx';
import { loadWordReferences } from 'ooxml-core/docx/ui';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { expect, test, vi } from 'vitest';
import { updateBodyFields } from './body-fields';
import { modelToDoc } from './model-adapter';

vi.mock(import('ooxml-core/docx/ui'), async (original) => {
	const real = await original();
	return { ...real, loadWordReferences: vi.fn(real.loadWordReferences) };
});

test('skips loading References for other fields and reports edits made during a load', async () => {
	const references = await loadWordReferences();
	vi.mocked(loadWordReferences).mockClear();
	const blank = createDocument();
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc: modelToDoc(blank) }),
	});
	const reportError = vi.fn();
	const host = { view: () => view, model: () => blank, reportError };
	try {
		await updateBodyFields(view, host, 'en');
		expect(loadWordReferences).not.toHaveBeenCalled();
		expect(reportError).toHaveBeenLastCalledWith(new Error('No fields needed updating.'));

		view.updateState(
			EditorState.create({
				doc: modelToDoc({
					...blank,
					blocks: [
						{
							type: 'paragraph',
							id: 'citation',
							runs: references.referenceFieldRuns('CITATION Doe20', [{ text: 'Cached' }]),
						},
					],
				}),
			}),
		);
		const updating = updateBodyFields(view, host, 'en');
		view.dispatch(view.state.tr.insertText('New text ', 1));
		await updating;
		expect(loadWordReferences).toHaveBeenCalledOnce();
		expect(reportError).toHaveBeenLastCalledWith(
			new Error('The document changed. Run Update Fields again.'),
		);
		expect(view.state.doc.textContent).toBe('New text Cached');
	} finally {
		view.destroy();
	}
});
