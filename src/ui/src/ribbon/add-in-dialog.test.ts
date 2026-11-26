import { afterEach, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import type { OfficeUiDialog } from '../dialog/dialog';
import { openOfficeAddIns } from './add-in-dialog';
import {
	setOfficeAddInCatalog,
	setOfficeAddInMarketplaceResolver,
	setOfficeAddInStorefront,
} from './add-in-settings';
import { setOfficeAddInPreferences, type OfficeAddInPreferences } from './add-in-preferences';

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

it('reviews resolved bytes before saving and restores and removes them through host preferences', async () => {
	const saved = new Map<string, Uint8Array>();
	const prefs = {
		read: async () => [...saved.values()],
		write: vi.fn<OfficeAddInPreferences['write']>(async (_host, id, bytes) => {
			if (bytes) saved.set(id, bytes);
			else saved.delete(id);
		}),
	};
	const make = () => {
		const host = document.body.appendChild(document.createElement('div'));
		setOfficeAddInPreferences(host, prefs);
		return host.attachShadow({ mode: 'open' });
	};
	const root = make();
	const resolver = vi.fn(async () => new TextEncoder().encode(xml));
	setOfficeAddInMarketplaceResolver(root, resolver);
	const open = (target: ShadowRoot) =>
		openOfficeAddIns(target, 'myAddIns', (label) => label, 'word', 'en');
	await open(root);
	const modal = getDialog(root);
	const input = modal?.querySelector<HTMLInputElement>('#marketplace-link');
	if (!input) throw new Error('Missing Marketplace field');
	await vi.waitFor(() => expect(input.disabled).toBe(false));
	input.value = 'https://appsource.microsoft.com/en-us/product/office/WA999999999';
	input.closest('form')?.dispatchEvent(new SubmitEvent('submit'));
	await vi.waitFor(() => expect(modal?.querySelector('#consent-title')).toBeTruthy());
	expect(resolver).toHaveBeenCalledWith('WA999999999', expect.any(AbortSignal));
	expect(prefs.write).not.toHaveBeenCalled();
	expect(modal?.querySelector('#pivot-2')).toBeNull();
	modal?.querySelector<HTMLButtonElement>('button.primary')?.click();
	await vi.waitFor(() => expect(getDialog(root)).toBeUndefined());
	expect(saved.size).toBe(1);
	root.host.remove();

	const reopened = make();
	await open(reopened);
	const mine = getDialog(reopened);
	await vi.waitFor(() =>
		expect(mine?.querySelector('summary')?.textContent).toBe('Add-in details'),
	);
	expect(mine?.querySelector('details')?.textContent).toContain('Lab Notes');
	mine?.querySelector<HTMLButtonElement>('button[aria-label^="Remove"]')?.click();
	await vi.waitFor(() => expect(saved.size).toBe(0));
});

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
	expect(dialog?.querySelector('[aria-selected="true"]')?.textContent).toBe('My Add-ins');
	expect(dialog?.querySelector('#pivot-2')).toBeNull();
	dialog?.querySelector<HTMLButtonElement>('#pivot-1')?.click();
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
	dialog?.querySelector<HTMLButtonElement>('#pivot-1')?.click();
	const search = dialog?.querySelector<HTMLInputElement>('input[type="search"]');
	if (!search) throw new Error('Missing search field');
	search.value = 'NOTES';
	search.dispatchEvent(new Event('input'));
	expect(dialog?.querySelector('li')?.textContent).toContain('Notes');
});

it('resolves the opener ribbon storefront when the dialog mounts', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const ribbon = editor.appendChild(document.createElement('div'));
	const control = ribbon.appendChild(document.createElement('div')).attachShadow({ mode: 'open' });
	const opener = control.appendChild(document.createElement('button'));
	setOfficeAddInCatalog(editor, [{ id: 'notes', label: 'Editor Notes', install: () => {} }]);
	setOfficeAddInStorefront(editor, 'off');
	opener.focus();
	const opening = openOfficeAddIns(editor, 'getAddIns', (text) => text, 'word', 'en');
	setOfficeAddInStorefront(ribbon, { kind: 'microsoft', label: 'Ribbon store' });
	await opening;
	const dialog = getDialog(editor);
	expect(dialog?.querySelector('a')?.textContent).toBe('Ribbon store');
	dialog?.querySelector<HTMLButtonElement>('#pivot-1')?.click();
	expect(dialog?.querySelector('li')?.textContent).toContain('Editor Notes');
	dialog?.close();
});

it('offers the Microsoft handoff and host catalog only when configured', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const open = () => openOfficeAddIns(editor, 'getAddIns', (text) => text, 'word', 'en');
	setOfficeAddInStorefront(editor, { kind: 'microsoft' });
	await open();
	expect(getDialog(editor)?.querySelector('[aria-selected="true"]')?.textContent).toBe('Store');
	const link = getDialog(editor)?.querySelector('a');
	expect(link?.textContent).toBe('Browse Microsoft Marketplace');
	expect(link?.hostname).toBe('marketplace.microsoft.com');
	expect(link?.rel).toBe('noopener noreferrer');
	getDialog(editor)?.close();
	await vi.waitFor(() => expect(getDialog(editor)).toBeUndefined());
	setOfficeAddInStorefront(editor, { kind: 'microsoft', url: 'javascript:alert(1)' });
	await open();
	expect(getDialog(editor)?.querySelector('a')).toBeNull();
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
	expect(dialog?.querySelector('[aria-selected="true"]')?.textContent).toBe('Store');
	expect(dialog?.querySelector('a')).toBeNull();
	expect(read).not.toHaveBeenCalled();
	dialog?.querySelector<HTMLButtonElement>('[aria-label="Install Notes"]')?.click();
	await vi.waitFor(() => expect(dialog?.querySelector('#consent-title')).toBeTruthy());
	expect(read).toHaveBeenCalledOnce();
	expect(dialog?.textContent).toContain('Read this document');
	dialog?.close();
});

it('explains why Visio cannot install an add-in from the Store', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const read = vi.fn(() => new File([xml], 'manifest.xml'));
	setOfficeAddInStorefront(editor, {
		kind: 'catalog',
		entries: [{ id: 'notes', label: 'Notes', read }],
	});
	await openOfficeAddIns(editor, 'getAddIns', (text) => text, 'visio', 'en');
	const dialog = getDialog(editor);
	const install = dialog?.querySelector<HTMLButtonElement>('[aria-label="Install Notes"]');
	const upload = dialog?.querySelector('#upload');
	const reason = upload?.getAttribute('aria-describedby');
	expect(install?.disabled).toBe(true);
	expect(install?.getAttribute('aria-describedby')).toBe(reason);
	expect(dialog?.querySelector(`[id="${reason}"]`)?.textContent).toBe(
		'Local add-in installation is unavailable in Visio.',
	);
	install?.click();
	expect(read).not.toHaveBeenCalled();
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
		description: { short: 'Reference notes.', full: 'Reference notes.' },
		developer: {
			name: 'Lab Notes',
			websiteUrl: 'https://example.test',
			privacyUrl: 'https://example.test/privacy',
			termsOfUseUrl: 'https://example.test/terms',
		},
		icons: { color: 'color.png', outline: 'outline.png' },
		accentColor: '#334455',
		localizationInfo: {
			defaultLanguageTag: 'en-us',
			additionalLanguages: [{ languageTag: 'fr-ca', file: 'fr-ca.json' }],
		},
		extensions: [{ requirements: { scopes: ['document'] } }],
	};
	const zip = await new JSZip()
		.file('manifest.json', JSON.stringify(manifest))
		.file('color.png', Uint8Array.of(1))
		.file('outline.png', Uint8Array.of(1))
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
	await vi.waitFor(() => expect(activeDialog(editor)).toBeUndefined());

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
	await vi.waitFor(() => expect(activeDialog(editor)).toBeUndefined());

	const newer = await open('getAddIns');
	await picker(newer, xml.replace('1.0.0', '1.0.1'));
	await vi.waitFor(() => expect(control(newer, 'Update')).toBeDefined());
	expect(newer.textContent).toContain('Update from 1.0.0 to 1.0.1');
	control(newer, 'Update')?.click();
	await vi.waitFor(() => expect(activeDialog(editor)).toBeUndefined());

	const updated = await open('myAddIns');
	await vi.waitFor(() => expect(updated.textContent).toContain('1.0.1'));
	expect(updated.textContent).not.toContain('1.0.0');
	control(updated, 'Remove')?.click();
	await vi.waitFor(() => expect(activeDialog(editor)).toBeUndefined());
	expect((await open('myAddIns')).textContent).toContain('No add-ins');
	activeDialog(editor)?.close();
});
