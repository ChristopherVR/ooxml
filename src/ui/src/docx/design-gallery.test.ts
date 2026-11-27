// @vitest-environment jsdom
import { afterEach, beforeAll, expect, it } from 'vitest';
import { loadThemePresets } from 'ooxml-core/drawingml';
import {
	createDocument,
	designSpacing,
	loadDocx,
	newShape,
	type WordDesignDefaults,
	type WordDesignScheme,
} from 'ooxml-core/docx';
import { DocxEditorElement, registerDocxEditor } from './component';

registerDocxEditor();
beforeAll(async () => {
	await Promise.all([import('./design-gallery'), loadThemePresets()]);
}, 60_000);
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
					shape: {
						...newShape('rect'),
						effectRef: {
							index: 3,
							color: { kind: 'scheme', value: 'accent1', transforms: [] },
						},
					},
				},
			},
		],
	});
	editor.documentModel = model;
	document.body.append(editor);
	const root = editor.shadowRoot!;
	// An open Design menu is searched first: the Equation tab has an Integral button too.
	const button = (name: string) =>
		[
			...root.querySelectorAll<HTMLButtonElement>('.dve-design-menu button'),
			...root.querySelectorAll('button'),
		].find((el) => el.getAttribute('aria-label') === name || el.textContent?.trim() === name);
	const open = async (label: string) => {
		const control = button(label);
		expect(control?.disabled).toBe(false);
		control?.click();
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
	button('Office')?.focus();
	button('Office')?.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }),
	);
	expect(root.activeElement).toBe(button('Integral'));
	button('Integral')?.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }),
	);
	expect(root.activeElement).toBe(button('Office'));
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
	await pick('Effects', 'Reflection');
	expect(root.querySelector('.dve-shape-drawing mask')).not.toBeNull();
	await pick('Effects', 'Top Shadow');
	expect(root.querySelector('.dve-shape-drawing feDropShadow')).not.toBeNull();
	editor.getToolbarProps().onUndo();
	expect(root.querySelector('.dve-shape-drawing mask')).not.toBeNull();
	await pick('Effects', 'Frosted Glass');
	expect(root.querySelector('.dve-shape-drawing feDropShadow')).not.toBeNull();
});

it('keeps named colour and font schemes across editors and supports rename, reuse and deletion', async () => {
	let schemes: WordDesignScheme[] = [];
	const preferences = {
		read: () => undefined,
		write: () => {},
		schemes: {
			read: () => schemes,
			write: (value: WordDesignScheme[]) => {
				schemes = value;
			},
		},
	};
	const setField = (root: ShadowRoot, label: string, value: string) => {
		const input = root.querySelector<HTMLInputElement>(
			`.dve-design-dialog input[aria-label="${label}"]`,
		);
		expect(input).not.toBeNull();
		if (!input) return;
		input.value = value;
		input.dispatchEvent(new Event('input', { bubbles: true }));
	};
	const save = (root: ShadowRoot) =>
		root.querySelector<HTMLButtonElement>('.dve-design-dialog .dve-dialog-primary')?.click();
	const first = mount();
	first.editor.wordDesignPreferences = preferences;
	await first.pick('Colors', 'Customize Colors');
	setField(first.root, 'Name', 'Research colors');
	setField(first.root, 'Accent 1', '#123456');
	save(first.root);
	await first.pick('Fonts', 'Customize Fonts');
	setField(first.root, 'Name', 'Research faces');
	setField(first.root, 'Heading font', 'Georgia');
	setField(first.root, 'Body font', 'Verdana');
	save(first.root);
	first.editor.remove();

	const second = mount();
	second.editor.wordDesignPreferences = preferences;
	await second.pick('Colors', 'Research colors');
	expect(second.editor.documentModel?.theme?.colors.accent1).toBe('123456');
	await second.pick('Fonts', 'Research faces');
	expect(second.editor.documentModel?.theme?.fonts.major.latin).toBe('Georgia');
	await second.pick('Fonts', 'Edit Research faces');
	setField(second.root, 'Name', 'Report faces');
	setField(second.root, 'Body font', 'Arial');
	save(second.root);
	expect(schemes).toMatchObject([
		{ name: 'Research colors' },
		{ name: 'Report faces', minorFont: 'Arial' },
	]);
	await second.open('Fonts');
	const manage = second.root.querySelector('[role="group"][aria-label="Manage custom schemes"]');
	expect(manage?.querySelector('[role="menuitem"][aria-label="Edit Report faces"]')).not.toBeNull();
	expect(
		second.root.querySelector('[role="group"][aria-label="Custom"] [role="menuitem"]'),
	).toBeNull();
	second.button('Edit Report faces')?.click();
	second.button('Delete')?.click();
	const confirmation = second.root.querySelector('[data-title="design.deleteSchemeTitle"]');
	expect(confirmation?.textContent).toContain('Report faces');
	expect(schemes).toHaveLength(2);
	const confirmButton = (panel: Element | null, label: string) =>
		[...(panel?.querySelectorAll('button') ?? [])].find((button) => button.textContent === label);
	confirmButton(confirmation, 'Cancel')?.click();
	expect(schemes).toHaveLength(2);
	second.button('Delete')?.click();
	confirmButton(
		second.root.querySelector('[data-title="design.deleteSchemeTitle"]'),
		'Delete',
	)?.click();
	expect(schemes).toMatchObject([{ name: 'Research colors' }]);
	expect((await loadDocx(await second.editor.saveBytes())).model.theme?.fonts.minor.latin).toBe(
		'Arial',
	);
});

it('provides keyboard navigation, custom colors and opt-in design defaults', async () => {
	const { editor, root, button, open, pick } = mount();
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
	await pick('Effects', 'Reflection');
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
	expect(stored?.effectName).toBe('Reflection');
	const blank = new DocxEditorElement();
	blank.wordDesignPreferences = editor.wordDesignPreferences;
	document.body.append(blank);
	expect(blank.documentModel?.theme?.colors.accent1).toBe('112233');
	expect(blank.documentModel?.blocks).toEqual(createDocument().blocks);
	expect((await loadDocx(await blank.saveBytes())).model.theme?.formatScheme?.name).toBe(
		'Reflection',
	);
	expect(blank.dirty).toBe(false);
	blank.documentModel = null;
	expect((await loadDocx(await blank.saveBytes())).model.theme?.formatScheme?.name).toBe(
		'Reflection',
	);
	blank.documentModel = null;
	blank.documentModel = createDocument();
	// A document given explicitly keeps its own design: it saves with the default Office theme.
	expect((await loadDocx(await blank.saveBytes())).model.theme?.formatScheme?.name).toBe('Office');
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
