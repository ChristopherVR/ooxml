import { afterEach, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import type { OfficeUiDialog } from '../dialog/dialog';
import { openOfficeAddIns, setOfficeAddInCatalog, setOfficeAddInStorefront } from './add-in-dialog';

const description = "<img src='https://example.test/description' />";
const descriptionPrefix = `${description}${'📝'.repeat(249 - description.length)}`;
const xml = `<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
	<Id>083a367e-842b-4d4f-ab91-cd5de849e495</Id><Version>1.0.0</Version>
	<ProviderName>Lab Notes</ProviderName><DefaultLocale>en</DefaultLocale>
	<DisplayName DefaultValue="&lt;img src='https://example.test/untrusted' /&gt;"/>
	<Description DefaultValue="${descriptionPrefix.replaceAll('&', '&amp;').replaceAll('<', '&lt;')}Reference notes."/>
	<DefaultSettings><SourceLocation DefaultValue="https://example.test/pane"/></DefaultSettings>
	<Permissions>ReadDocument</Permissions>
</OfficeApp>`;

afterEach(() => document.body.replaceChildren());

const getDialog = (root: ShadowRoot) =>
	root
		.querySelector('office-ui-add-ins-dialog')
		?.shadowRoot?.querySelector<OfficeUiDialog>('office-ui-dialog[open]');

it('opens at the editor root, filters the catalog and restores the opener on Escape', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const opener = editor.appendChild(document.createElement('button'));
	opener.focus();
	setOfficeAddInCatalog(editor, [
		{ id: 'host', label: 'Host add-in', install: () => {} },
		{ id: 'other', label: 'Other add-in', install: () => {} },
	]);
	await openOfficeAddIns(editor, 'getAddIns', (text) => text, 'word', 'en');
	const dialog = getDialog(editor);
	expect(dialog?.heading).toBe('Office Add-ins');
	const search = dialog?.querySelector<HTMLInputElement>('input[type="search"]');
	if (!search) throw new Error('Missing search field');
	expect(dialog?.getRootNode()).toBe(editor.querySelector('office-ui-add-ins-dialog')?.shadowRoot);
	expect(dialog?.querySelector('a')).toBeNull();
	expect(dialog?.querySelectorAll('li')).toHaveLength(2);
	search.value = 'HOST';
	search.dispatchEvent(new Event('input'));
	expect(dialog?.querySelectorAll('li')).toHaveLength(1);
	expect(dialog?.querySelector('li')?.textContent).toContain('Host add-in');
	search.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'Escape', composed: true, bubbles: true }),
	);
	await vi.waitFor(() => expect(getDialog(editor)).toBeUndefined());
	expect(editor.activeElement).toBe(opener);
});

it('uses a normalized locale for both labels and catalog search', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	setOfficeAddInCatalog(editor, [{ id: 'notes', label: 'Notes', install: () => {} }]);
	await openOfficeAddIns(editor, 'getAddIns', (text) => text, 'word', 'zh_CN');
	const dialog = getDialog(editor);
	expect(dialog?.heading).toBe('Office 加载项');
	const search = dialog?.querySelector<HTMLInputElement>('input[type="search"]');
	if (!search) throw new Error('Missing search field');
	search.value = 'NOTES';
	search.dispatchEvent(new Event('input'));
	expect(dialog?.querySelector('li')?.textContent).toContain('Notes');
});

it('offers only the configured storefront and validates BYO bytes before installing', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const open = () => openOfficeAddIns(editor, 'getAddIns', (text) => text, 'word', 'en');
	setOfficeAddInStorefront(editor, { kind: 'microsoft', url: 'javascript:alert(1)' });
	await open();
	expect(getDialog(editor)?.querySelector('a')).toBeNull();
	getDialog(editor)?.close();
	await vi.waitFor(() => expect(getDialog(editor)).toBeUndefined());
	setOfficeAddInStorefront(editor, { kind: 'microsoft' });
	await open();
	const link = getDialog(editor)?.querySelector('a');
	expect(link?.hostname).toBe('marketplace.microsoft.com');
	expect(link?.rel).toBe('noopener noreferrer');
	getDialog(editor)?.close();
	await vi.waitFor(() => expect(getDialog(editor)).toBeUndefined());
	const read = vi.fn(function (this: { id: string }, signal: AbortSignal) {
		expect(this.id).toBe('notes');
		expect(signal.aborted).toBe(false);
		return new File([xml], 'manifest.xml');
	});
	setOfficeAddInStorefront(editor, {
		kind: 'catalog',
		entries: [{ id: 'notes', label: 'Notes', read }],
	});
	await open();
	const dialog = getDialog(editor);
	expect(dialog?.querySelector('a')).toBeNull();
	expect(read).not.toHaveBeenCalled();
	dialog?.querySelector<HTMLButtonElement>('[aria-label="Install Notes"]')?.click();
	await vi.waitFor(() => expect(dialog?.querySelector('#consent-title')).toBeTruthy());
	expect(read).toHaveBeenCalledOnce();
	expect(dialog?.textContent).toContain('Read this document');
	dialog?.close();
});

it('uses a regional host locale for unified manifest translations', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const manifest = {
		manifestVersion: '1.25',
		id: '083a367e-842b-4d4f-ab91-cd5de849e495',
		version: '1.0.0',
		name: { short: 'Notes' },
		description: { short: 'Reference notes.' },
		developer: { name: 'Lab Notes' },
		localizationInfo: {
			defaultLanguageTag: 'en-us',
			additionalLanguages: [{ languageTag: 'fr-ca', file: 'fr-ca.json' }],
		},
		extensions: [{ requirements: { scopes: ['document'] } }],
	};
	const zip = await new JSZip()
		.file('manifest.json', JSON.stringify(manifest))
		.file('fr-ca.json', JSON.stringify({ 'name.short': 'Notes du Québec' }))
		.generateAsync({ type: 'arraybuffer' });
	await openOfficeAddIns(editor, 'getAddIns', (text) => text, 'word', 'fr-CA');
	const dialog = getDialog(editor);
	expect(dialog?.heading).toBe('Compléments Office');
	const input = dialog?.querySelector<HTMLInputElement>('input[type=file]');
	if (!input) throw new Error('Missing file input');
	Object.defineProperty(input, 'files', { value: [new File([zip], 'notes.zip')] });
	input.dispatchEvent(new Event('change'));
	await vi.waitFor(() =>
		expect(dialog?.querySelector('.tile-name')?.textContent).toBe('Notes du Québec (1.0.0)'),
	);
	dialog?.close();
});

it('requires consent, keeps the prior install through a failed update and isolates editors', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const other = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const activeDialog = getDialog;
	const open = async (mode: 'getAddIns' | 'myAddIns', root = editor) => {
		await openOfficeAddIns(root, mode, (text) => text, 'word', 'en');
		return vi.waitFor(() => {
			const dialog = activeDialog(root);
			if (!dialog) throw new Error(`Missing ${mode} dialog`);
			return dialog;
		});
	};
	const picker = async (root: Awaited<ReturnType<typeof open>>, source: string) => {
		const input = await vi.waitFor(() => {
			const fileInput = root.querySelector<HTMLInputElement>('input[type=file]');
			if (!fileInput) throw new Error(`Missing file input. Dialog content: ${root.textContent}`);
			return fileInput;
		});
		const file = new File([source], 'manifest.xml');
		Object.defineProperty(input, 'files', { configurable: true, value: [file] });
		input.dispatchEvent(new Event('change'));
	};
	const control = (root: ParentNode, label: string) =>
		[...root.querySelectorAll('button')].find((button) => button.textContent === label);

	const abandoned = await open('getAddIns');
	await picker(abandoned, xml);
	await vi.waitFor(() => expect(control(abandoned, 'Install')).toBeDefined());
	const consent = abandoned.querySelector('section[aria-labelledby="consent-title"]');
	expect(consent?.textContent).toContain('This add-in requests:');
	expect(consent?.textContent).toContain('Read this document');
	expect(consent?.textContent).toContain('https://example.test');
	expect(consent?.textContent).toContain('Lab Notes');
	expect(consent?.textContent).not.toContain('Document.Read.User');
	expect([...abandoned.querySelectorAll('p')].map((item) => item.textContent)).toContain(
		`${descriptionPrefix}…`,
	);
	expect(abandoned.querySelector('img, iframe')).toBeNull();
	control(abandoned, 'Cancel')?.click();
	expect(abandoned.querySelector('#consent-title')).toBeNull();
	expect(abandoned.querySelector('#upload')).toHaveProperty('disabled', false);
	abandoned.close();
	expect((await open('myAddIns')).textContent).toContain('No add-ins');
	activeDialog(editor)?.close();

	const get = await open('getAddIns');
	await picker(get, xml);
	const install = await vi.waitFor(() => {
		const button = control(get, 'Install');
		if (!button) throw new Error('Missing consent control');
		return button;
	});
	install.click();

	const invalid = await open('getAddIns');
	await picker(invalid, xml.replace('https://example.test/pane', 'javascript:alert(1)'));
	await vi.waitFor(() => expect(invalid.textContent).toContain('unsafe URL or origin'));
	expect(control(invalid, 'Update')).toBeUndefined();
	invalid.close();
	const mine = await open('myAddIns');
	await vi.waitFor(() => expect(mine.textContent).toContain('1.0.0'));
	expect(mine.textContent).toContain('Running add-ins is unavailable');
	expect((await open('myAddIns', other)).textContent).toContain('No add-ins');
	activeDialog(other)?.close();
	mine.close();

	const sameVersion = await open('getAddIns');
	await picker(sameVersion, xml.replace('Lab Notes', 'New Publisher'));
	await vi.waitFor(() => expect(control(sameVersion, 'Update')).toBeDefined());
	expect(sameVersion.textContent).toContain('Update this add-in and keep version 1.0.0');
	control(sameVersion, 'Update')?.click();

	const newer = await open('getAddIns');
	await picker(newer, xml.replace('1.0.0', '1.0.1'));
	await vi.waitFor(() => expect(control(newer, 'Update')).toBeDefined());
	expect(newer.textContent).toContain('Update from 1.0.0 to 1.0.1');
	control(newer, 'Update')?.click();

	const updated = await open('myAddIns');
	await vi.waitFor(() => expect(updated.textContent).toContain('1.0.1'));
	expect(updated.textContent).not.toContain('1.0.0');
	control(updated, 'Remove')?.click();
	await vi.waitFor(() => expect(activeDialog(editor)).toBeUndefined());
	expect((await open('myAddIns')).textContent).toContain('No add-ins');
	activeDialog(editor)?.close();
});
