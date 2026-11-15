// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, signedTwips, type Table, type TableStyleCatalog } from 'ooxml-core/docx';
import { EditorState, TextSelection } from 'prosemirror-state';
import { modelToDoc } from './model-adapter';
import { tablePropertiesContext } from './table-properties';

function contextFor(table: Partial<Table>, tableStyles?: TableStyleCatalog) {
	const model = createDocument();
	model.blocks = [
		{
			type: 'table',
			id: 'table',
			...table,
			rows: [[{ paragraphs: [{ type: 'paragraph', id: 'p', runs: [{ text: 'cell' }] }] }]],
		},
	];
	if (tableStyles) model.tableStyles = tableStyles;
	const doc = modelToDoc(model);
	const state = EditorState.create({ doc, selection: TextSelection.create(doc, 4) });
	const context = tablePropertiesContext(state);
	if (!context) throw new Error('no table context');
	return context;
}

const catalog: TableStyleCatalog = {
	warnings: [],
	styles: {
		TableNormal: {
			id: 'TableNormal',
			conditional: {},
			isDefault: true,
			cellMargins: { left: signedTwips(50), right: signedTwips(50) },
		},
		Base: { id: 'Base', conditional: {}, cellMargins: { left: signedTwips(200) } },
		Child: {
			id: 'Child',
			conditional: {},
			basedOn: 'Base',
			cellMargins: { right: signedTwips(300) },
		},
	},
};

describe('table dialogs: inherited cell margins', () => {
	it('shows the margins the table takes from its style chain, as layout does', () => {
		const context = contextFor({ style: 'Child' }, catalog);
		expect(context.margins).toEqual({});
		expect(context.inheritedMargins.left).toBe(200);
		expect(context.inheritedMargins.right).toBe(300);
	});

	it('falls back to the default table style', () => {
		expect(contextFor({}, catalog).inheritedMargins.left).toBe(50);
	});

	it("keeps Word's built-in 108 twips when the document has no styles part", () => {
		const context = contextFor({});
		expect(context.inheritedMargins.left).toBe(108);
		expect(context.inheritedMargins.top).toBe(0);
	});
});
