// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import {
	createDocument,
	designSpacing,
	loadDocx,
	newShape,
	type WordDesignDefaults,
} from 'ooxml-core/docx';
import { DocxEditorElement, registerDocxEditor } from './component';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());

function mount() {
	const editor = new DocxEditorElement();
	const model = createDocument();
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p1',
			style: 'Heading1',
			runs: [{ text: 'Themed' }, { text: 'Direct', fontFamily: 'Arial', color: '#123456' }],
		},
	];
	model.blocks.push({
		type: 'paragraph',
		id: 'shape',
		runs: [
			{
				text: '',
				image: {
					relId: '',
					partName: '',
					contentType: 'application/octet-stream',
					widthPx: 120,
					heightPx: 60,
					shape: newShape('rect'),
				},
			},
		],
	});
	editor.documentModel = model;
	document.body.append(editor);
	const root = editor.shadowRoot!;
	const button = (name: string) =>
		[...root.querySelectorAll('button')].find(
			(el) => el.getAttribute('aria-label') === name || el.textContent?.trim() === name,
		);
	const open = async (label: string) => {
		button(label)?.click();
		await expect.poll(() => root.querySelector('.dve-design-menu')).not.toBeNull();
	};
	const pick = async (menu: string, name: string) => {
		await open(menu);
		const item = button(name);
		expect(item).toBeDefined();
		item?.click();
	};
	root
		.querySelector<HTMLButtonElement>('office-ui-ribbon')
		?.shadowRoot?.querySelector<HTMLButtonElement>('[data-tab="design"]')
		?.click();
	return { editor, root, button, pick, open };
}

it('previews gallery items, changes theme formatting live, and undoes one choice at a time', async () => {
	const { editor, root, button, pick, open } = mount();
	await open('Themes');
	expect(root.querySelectorAll('.dve-design-menu .dve-design-swatches').length).toBeGreaterThan(2);
	button('Facet')?.click();
	expect(editor.documentModel?.theme?.name).toBe('Facet');
	expect(root.querySelector('.ProseMirror p')?.innerHTML).toContain('Trebuchet MS');
	expect(root.querySelector('.ProseMirror p')?.innerHTML).toMatch(/rgb\(18, 52, 86\)|#123456/);
	await pick('Paragraph spacing', 'Double');
	expect(editor.documentModel && designSpacing(editor.documentModel).line).toBe(480);
	editor.getToolbarProps().onUndo();
	expect(editor.documentModel?.theme?.name).toBe('Facet');
	expect(editor.documentModel && designSpacing(editor.documentModel).line).toBe(259);
	editor.getToolbarProps().onUndo();
	expect(editor.documentModel?.theme).toBeUndefined();
	editor.getToolbarProps().onRedo();
	expect(editor.documentModel?.theme?.name).toBe('Facet');
	expect((await loadDocx(await editor.saveBytes())).model.theme?.name).toBe('Facet');
	await pick('Effects', 'No Shadow');
	expect(root.querySelector('.dve-shape-drawing feDropShadow')).toBeNull();
	await pick('Effects', 'Strong Shadow');
	expect(root.querySelector('.dve-shape-drawing feDropShadow')).not.toBeNull();
	editor.getToolbarProps().onUndo();
	expect(root.querySelector('.dve-shape-drawing feDropShadow')).toBeNull();
});

it('provides keyboard gallery navigation, a custom color form and opt-in defaults', async () => {
	const { editor, root, button, open } = mount();
	expect(button('Set as Default')?.disabled).toBe(true);
	await open('Colors');
	const current = root.querySelector('[role="menuitemradio"][aria-checked="true"]');
	expect(current).not.toBeNull();
	current?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
	expect(root.activeElement?.textContent).toBe('Customize Colors');
	button('Customize Colors')?.click();
	const accent = root.querySelector<HTMLInputElement>('input[aria-label="Accent 1"]');
	expect(accent).not.toBeNull();
	if (accent) accent.value = '#112233';
	root.querySelector<HTMLButtonElement>('.dve-design-dialog .dve-dialog-primary')?.click();
	expect(editor.documentModel?.theme?.colors.accent1).toBe('112233');
	let stored: WordDesignDefaults | undefined;
	editor.wordDesignPreferences = {
		read: () => stored,
		write: (value) => {
			stored = value;
		},
	};
	button('Set as Default')?.click();
	await expect.poll(() => root.querySelector('.dve-design-dialog')).not.toBeNull();
	root.querySelector<HTMLButtonElement>('.dve-design-dialog .dve-dialog-primary')?.click();
	expect(stored?.theme.colors.accent1).toBe('112233');
	const blank = new DocxEditorElement();
	blank.wordDesignPreferences = editor.wordDesignPreferences;
	document.body.append(blank);
	expect(blank.documentModel?.theme?.colors.accent1).toBe('112233');
	expect(blank.documentModel?.blocks).toEqual(createDocument().blocks);
	editor.wordDesignPreferences = undefined;
	expect(button('Set as Default')?.disabled).toBe(true);
});
it('saves custom font pairs and exact paragraph spacing from their dialogs', async () => {
	const { editor, root, button, open } = mount();
	const writeField = (name: string, value: string) => {
		const input = root.querySelector<HTMLInputElement>(
			`.dve-design-dialog input[aria-label="${name}"]`,
		);
		expect(input).not.toBeNull();
		if (input) {
			input.value = value;
			input.dispatchEvent(new Event('input', { bubbles: true }));
		}
	};
	const confirm = () =>
		root.querySelector<HTMLButtonElement>('.dve-design-dialog .dve-dialog-primary')?.click();
	await open('Fonts');
	button('Customize Fonts')?.click();
	writeField('Name', 'Reading');
	writeField('Heading font', 'Georgia');
	writeField('Body font', 'Verdana');
	confirm();
	await open('Paragraph spacing');
	button('Custom Paragraph Spacing')?.click();
	const rule = root.querySelector<HTMLSelectElement>('.dve-design-dialog select');
	expect(rule).not.toBeNull();
	if (rule) {
		rule.value = 'exact';
		rule.dispatchEvent(new Event('change', { bubbles: true }));
	}
	writeField('Line height (pt)', '14');
	writeField('Before (pt)', '2');
	writeField('After (pt)', '7');
	confirm();
	const { model } = await loadDocx(await editor.saveBytes());
	expect(model.theme).toMatchObject({
		fontSchemeName: 'Reading',
		fonts: { major: { latin: 'Georgia' }, minor: { latin: 'Verdana' } },
	});
	expect(designSpacing(model)).toEqual({ before: 40, after: 140, line: 280, rule: 'exact' });
});
