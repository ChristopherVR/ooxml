// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { afterEach, describe, expect, it } from 'vitest';
import { selectionInTable, syncContextualTabs } from './contextual-tabs';
import { createRibbon } from './ribbon';
import { schema } from './schema';
import { ribbonTab, ribbonTabs } from './test-support';

function tableState() {
	const cell = schema.nodes.tableCell!.create(
		{},
		schema.nodes.paragraph!.create({ id: 'c' }, schema.text('cell')),
	);
	const table = schema.nodes.table!.create({ id: 't' }, schema.nodes.tableRow!.create({}, cell));
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create({ id: 'a' }, schema.text('before')),
		table,
	]);
	return EditorState.create({ doc, schema });
}

const mount = () => {
	const ribbon = createRibbon();
	document.body.append(ribbon);
	return ribbon;
};
/** Waits for the shared ribbon to catch up with its panels (it watches their attributes). */
const settle = async (ribbon: HTMLElement) => {
	await new Promise((resolve) => setTimeout(resolve, 0));
	await (ribbon as HTMLElement & { updateComplete: Promise<unknown> }).updateComplete;
};
/** The keys of the tabs the ribbon shows. */
const keys = async (ribbon: HTMLElement) => {
	await settle(ribbon);
	return (await ribbonTabs(ribbon)).map((tab) => tab.dataset.tab);
};
const inTable = (state: EditorState) =>
	state.apply(state.tr.setSelection(TextSelection.create(state.doc, state.doc.content.size - 4)));
const outside = (state: EditorState) =>
	state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));

afterEach(() => (document.body.innerHTML = ''));

describe('contextual Table tab', () => {
	it('has no tab in a new ribbon, and is tinted as a contextual tab once shown', async () => {
		const ribbon = mount();
		expect(await keys(ribbon)).not.toContain('table');
		syncContextualTabs(ribbon, inTable(tableState()));
		expect(await keys(ribbon)).toContain('table');
		expect((await ribbonTab(ribbon, 'table')).hasAttribute('data-contextual')).toBe(true);
	});

	it('appears while the selection is in a table and leaves when it moves out', async () => {
		const ribbon = mount();
		let state = inTable(tableState());
		expect(selectionInTable(state)).toBe(true);
		syncContextualTabs(ribbon, state);
		expect(await keys(ribbon)).toContain('table');
		state = outside(state);
		syncContextualTabs(ribbon, state);
		expect(await keys(ribbon)).not.toContain('table');
	});

	it('returns to Home when the table tab was open as the selection leaves', async () => {
		const ribbon = mount();
		let state = inTable(tableState());
		syncContextualTabs(ribbon, state);
		await settle(ribbon);
		(await ribbonTab(ribbon, 'table')).click();
		await settle(ribbon);
		expect(ribbon.querySelector('#dve-panel-table')!.hasAttribute('hidden')).toBe(false);
		state = outside(state);
		syncContextualTabs(ribbon, state);
		await settle(ribbon);
		expect((await ribbonTab(ribbon, 'home')).getAttribute('aria-selected')).toBe('true');
		expect(ribbon.querySelector('#dve-panel-table')!.hasAttribute('hidden')).toBe(true);
	});

	it('skips the hidden tab when arrowing through the tab strip', async () => {
		const ribbon = mount();
		const view = await ribbonTab(ribbon, 'view');
		view.focus();
		view.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true, composed: true }),
		);
		await settle(ribbon);
		expect((await ribbonTab(ribbon, 'home')).getAttribute('aria-selected')).toBe('true');
	});
});
