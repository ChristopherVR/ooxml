// @vitest-environment jsdom
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import { afterEach, describe, expect, it } from 'vitest';
import { inspectDocument, loadDocx, type DocumentModel } from 'ooxml-core/docx';
import { trackChangesPlugin } from 'ooxml-core/docx/ui';
import { inspectFixture } from '../../../core/docx/test-support/inspect-fixture';
import { docToModel, modelToDoc } from './model-adapter';
import { schema } from './schema';
import { inspectorActions } from './document-inspector-apply';
import { createDocumentInspectorDialog } from './document-inspector-dialog';

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});

/** An editor whose dispatch keeps `model` in step with the document, as the element does. */
async function editor(trackChanges = false) {
	const { model: loaded } = await loadDocx(await inspectFixture());
	const state = { model: { ...loaded, trackChanges } as DocumentModel };
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({
			doc: modelToDoc(state.model),
			schema,
			plugins: [
				history(),
				trackChangesPlugin(
					() => 'Editor',
					() => trackChanges,
				),
			],
		}),
		dispatchTransaction(tr) {
			const previous = view.state.doc;
			view.updateState(view.state.apply(tr));
			state.model = docToModel(view.state.doc, state.model, previous);
		},
	});
	views.push(view);
	const actions = inspectorActions({
		view: () => view,
		model: () => state.model,
		canEdit: () => true,
		reportError: (cause) => {
			throw cause;
		},
	});
	return { view, state, actions, original: state.model };
}

const found = (model: DocumentModel) =>
	inspectDocument(model)
		.filter((entry) => entry.findings.length)
		.map((entry) => entry.category);

describe('Document Inspector removal', () => {
	it.each(['comments', 'properties', 'customXml', 'headersFooters', 'hiddenText'] as const)(
		'removes %s as one undoable step',
		async (category) => {
			const { view, state, actions, original } = await editor();
			expect(found(state.model)).toContain(category);
			expect(actions.remove(category)).toBe(true);
			expect(found(state.model)).not.toContain(category);
			expect(undo(view.state, view.dispatch)).toBe(true);
			expect(inspectDocument(state.model)).toEqual(inspectDocument(original));
			expect(state.model.properties).toEqual(original.properties);
			expect(state.model.packageInformation).toEqual(original.packageInformation);
		},
	);

	it('does not record the removal as a tracked change', async () => {
		const { state, actions } = await editor(true);
		actions.remove('hiddenText');
		actions.remove('headersFooters');
		expect(found(state.model)).toEqual(['comments', 'properties', 'customXml']);
		expect(inspectDocument(state.model, ['comments'])[0]!.findings).toEqual([
			{ kind: 'comments', count: 1 },
			{ kind: 'revisions', count: 1 },
		]);
	});

	it('keeps a File > Info property edit through later edits', async () => {
		const { view, state } = await editor();
		state.model.properties = { ...state.model.properties, title: 'Renamed' };
		view.dispatch(view.state.tr.insertText('x', 2));
		expect(state.model.properties?.title).toBe('Renamed');
	});

	it('refuses removal under tracked-changes protection', async () => {
		const { state, actions } = await editor();
		state.model = { ...state.model, protection: { edit: 'trackedChanges', enforced: true } };
		expect(actions.blocked('comments')).toBe('protected');
		expect(actions.remove('comments')).toBe(false);
		expect(actions.blocked('hiddenText')).toBeUndefined();
	});
});

describe('Document Inspector dialog', () => {
	it('inspects, lists what each category found and removes it', async () => {
		const { state, actions } = await editor();
		const dialog = createDocumentInspectorDialog({
			model: () => state.model,
			...actions,
			restoreFocus: () => {},
		});
		document.body.append(dialog.element);
		dialog.open();
		const rows = () => [...dialog.element.querySelectorAll<HTMLElement>('.dve-inspector-row')];
		expect(rows().map((row) => row.querySelector('strong')?.textContent)).toEqual([
			'Comments, Revisions, and Versions',
			'Document Properties and Personal Information',
			'Custom XML Data',
			'Headers, Footers, and Watermarks',
			'Hidden Text',
		]);
		dialog.element.querySelector<HTMLInputElement>('input[data-category="customXml"]')!.click();
		const button = (label: string) =>
			[...dialog.element.querySelectorAll('button')].find((item) => item.textContent === label)!;
		button('Inspect').click();
		expect(rows().map((row) => row.dataset.category)).toEqual([
			'comments',
			'properties',
			'headersFooters',
			'hiddenText',
		]);
		const hidden = rows().find((row) => row.dataset.category === 'hiddenText')!;
		expect(hidden.dataset.state).toBe('found');
		expect(hidden.textContent).toContain('Hidden text');
		hidden.querySelector('button')!.click();
		const cleared = rows().find((row) => row.dataset.category === 'hiddenText')!;
		expect(cleared.dataset.state).toBe('clear');
		expect(cleared.textContent).toContain('All items were removed.');
		expect(found(state.model)).not.toContain('hiddenText');
		button('Reinspect').click();
		expect(button('Inspect')).toBeDefined();
		button('Close').click();
		expect(dialog.isOpen).toBe(false);
	});

	it('translates its text', async () => {
		const { state, actions } = await editor();
		const dialog = createDocumentInspectorDialog({
			model: () => state.model,
			...actions,
			restoreFocus: () => {},
		});
		document.body.append(dialog.element);
		dialog.setLocale('fr');
		dialog.open();
		expect(dialog.element.textContent).toContain('Texte masqué');
		expect(dialog.element.textContent).toContain('Inspecter');
	});
});
