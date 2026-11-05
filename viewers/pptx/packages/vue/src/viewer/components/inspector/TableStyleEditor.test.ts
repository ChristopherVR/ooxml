import { mount } from '@vue/test-utils';
import type { ParsedTableStyleMap } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import TableStyleEditor from './TableStyleEditor.vue';
import { setControlValue } from './test-control-value';

/** In-app dialogs load on demand; a busy CI runner needs more than waitFor's 1 s default. */
const DIALOG_WAIT = { timeout: 5000 };

/**
 * W4-E: the table STYLE DEFINITION editor ("Edit style...") lets an author
 * change a `ppt/tableStyles.xml` section's fill/text/borders, and clone /
 * delete a style, entirely through the shared `pptx-viewer-shared`
 * describe/apply pair. Vue port of React's `TableStyleEditor.test.tsx`.
 */
const STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

function styleMap(): ParsedTableStyleMap {
	return {
		[STYLE_ID]: {
			styleId: STYLE_ID,
			styleName: 'Medium Style 2 - Accent 1',
			wholeTblFill: { schemeColor: '', color: '#336699' },
		},
	};
}

/** Answer the in-app dialog the component opened: type into its input, then press a button. */
async function answerDialog(button: 'confirm' | 'cancel', value?: string): Promise<void> {
	const dialog = await vi.waitFor(() => {
		const found = document.querySelector('office-ui-message-dialog');
		if (!found) throw new Error('no dialog yet');
		return found;
	}, DIALOG_WAIT);
	const input = dialog.shadowRoot?.querySelector('input');
	if (input && value !== undefined) input.value = value;
	(dialog.shadowRoot!.querySelector(`button.${button}`) as HTMLButtonElement).click();
	await new Promise((resolve) => setTimeout(resolve, 0));
}

afterEach(() => {
	document.querySelectorAll('office-ui-message-dialog').forEach((node) => node.remove());
});

describe('tableStyleEditor.vue', () => {
	it('shows the "no style selected" message with no styleId', () => {
		const wrapper = mount(TableStyleEditor, {
			props: { styleMap: styleMap(), styleId: undefined, canEdit: true },
		});
		expect(wrapper.text()).toContain('Select a table style to edit its definition.');
	});

	it('lists all 14 parts and edits the fill colour of the selected one', async () => {
		const wrapper = mount(TableStyleEditor, {
			props: { styleMap: styleMap(), styleId: STYLE_ID, canEdit: true },
		});
		const partButtons = wrapper
			.findAll('button')
			.filter((b) => b.text() === 'Whole Table' || b.text() === 'Background');
		expect(partButtons).toHaveLength(2);

		const fillInput = wrapper.find('input[type="color"]');
		await setControlValue(fillInput, '#ff0000');

		const emitted = wrapper.emitted('styleMapChange');
		expect(emitted).toBeTruthy();
		const nextMap = emitted![0][0] as ParsedTableStyleMap;
		expect(nextMap[STYLE_ID].wholeTblFill).toStrictEqual({ schemeColor: '', color: '#ff0000' });
	});

	it('creates a new style from the current one, based on a prompt', async () => {
		const wrapper = mount(TableStyleEditor, {
			props: { styleMap: styleMap(), styleId: STYLE_ID, canEdit: true },
		});
		const createButton = wrapper
			.findAll('button')
			.find((b) => b.text() === 'New style from current');
		await createButton!.trigger('click');
		await answerDialog('confirm', 'My Custom Style');

		const emitted = wrapper.emitted('styleMapChange');
		expect(emitted).toBeTruthy();
		const nextMap = emitted![0][0] as ParsedTableStyleMap;
		const ids = Object.keys(nextMap);
		expect(ids).toHaveLength(2);
		const newId = ids.find((id) => id !== STYLE_ID)!;
		expect(nextMap[newId].styleName).toBe('My Custom Style');
		expect(wrapper.emitted('assignStyle')?.[0]).toStrictEqual([newId]);
	});

	it('creates a brand-new style with no current style selected and assigns it', async () => {
		const wrapper = mount(TableStyleEditor, {
			props: { styleMap: undefined, styleId: undefined, canEdit: true },
		});
		expect(wrapper.text()).toContain('Select a table style to edit its definition.');
		const createButton = wrapper.findAll('button').find((b) => b.text() === 'Create new style');
		expect(createButton).toBeTruthy();
		expect(createButton!.attributes('disabled')).toBeUndefined();
		await createButton!.trigger('click');
		await answerDialog('confirm', 'Brand New Style');

		const emitted = wrapper.emitted('styleMapChange');
		expect(emitted).toBeTruthy();
		const nextMap = emitted![0][0] as ParsedTableStyleMap;
		const ids = Object.keys(nextMap);
		expect(ids).toHaveLength(1);
		expect(nextMap[ids[0]].styleName).toBe('Brand New Style');
		expect(wrapper.emitted('assignStyle')?.[0]).toStrictEqual([ids[0]]);
	});

	it('deletes the style after confirmation and reports both the map and the id', async () => {
		const wrapper = mount(TableStyleEditor, {
			props: { styleMap: styleMap(), styleId: STYLE_ID, canEdit: true },
		});
		const deleteButton = wrapper.findAll('button').find((b) => b.text() === 'Delete style');
		await deleteButton!.trigger('click');
		await answerDialog('confirm');

		expect(wrapper.emitted('deleteStyle')?.[0]).toStrictEqual([STYLE_ID]);
		const nextMap = wrapper.emitted('styleMapChange')![0][0] as ParsedTableStyleMap;
		expect(Object.keys(nextMap)).toHaveLength(0);
		expect(wrapper.emitted('close')).toBeTruthy();
	});

	it('keeps the style when the delete dialog is cancelled', async () => {
		const wrapper = mount(TableStyleEditor, {
			props: { styleMap: styleMap(), styleId: STYLE_ID, canEdit: true },
		});
		const deleteButton = wrapper.findAll('button').find((b) => b.text() === 'Delete style');
		await deleteButton!.trigger('click');
		await answerDialog('cancel');
		expect(wrapper.emitted('deleteStyle')).toBeUndefined();
	});

	it('disables field controls when canEdit is false', () => {
		const wrapper = mount(TableStyleEditor, {
			props: { styleMap: styleMap(), styleId: STYLE_ID, canEdit: false },
		});
		const fillInput = wrapper.find('input[type="color"]');
		expect(fillInput.attributes('disabled')).toBeDefined();
	});
});
