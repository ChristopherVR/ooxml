import { afterEach, expect, it, vi } from 'vitest';
import type { OfficeUiDialog } from '../dialog/dialog';
import { openOfficeAddIns, setOfficeAddInCatalog } from './add-in-dialog';
import './add-in-local-dialog';

const description = "<img src='https://example.test/description' />";
const descriptionPrefix = `${description}${'📝'.repeat(249 - description.length)}`;
const xml = `<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
	<Id>083a367e-842b-4d4f-ab91-cd5de849e495</Id><Version>1.0.0</Version>
	<ProviderName>Lab Notes</ProviderName><DefaultLocale>en</DefaultLocale>
	<DisplayName DefaultValue="&lt;img src='https://example.test/untrusted' /&gt;"/>
	<Description DefaultValue="${descriptionPrefix.replace('<', '&lt;')}Reference notes."/>
	<DefaultSettings><SourceLocation DefaultValue="https://example.test/pane"/></DefaultSettings>
	<Permissions>ReadDocument</Permissions>
</OfficeApp>`;

afterEach(() => document.body.replaceChildren());

it('opens while the local picker loads and focuses it ahead of the host catalog', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	setOfficeAddInCatalog(editor, [{ id: 'host', label: 'Host add-in', install: () => {} }]);
	openOfficeAddIns(editor, 'getAddIns', (text) => text, 'word', 'fr');
	const dialog = editor.querySelector<OfficeUiDialog>('office-ui-dialog[open]');
	expect(dialog?.querySelector('[role="status"]')?.textContent).toBe('Chargement des compléments…');
	await vi.waitFor(() => {
		const input = dialog?.querySelector<HTMLInputElement>('input[type=file]');
		expect(input).toBeInstanceOf(HTMLInputElement);
		expect(editor.activeElement).toBe(input);
	});
	expect(dialog?.textContent).not.toContain('Chargement des compléments…');
	expect(dialog?.querySelector('[aria-label="Install Host add-in"]')).toBeInstanceOf(
		HTMLButtonElement,
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
	const activeDialog = (root: ShadowRoot) =>
		root.querySelector<OfficeUiDialog>('office-ui-dialog[open]');
	const open = async (mode: 'getAddIns' | 'myAddIns', root = editor) => {
		openOfficeAddIns(root, mode, (text) => text, 'word', 'en');
		return vi.waitFor(() => {
			const dialog = activeDialog(root);
			if (!dialog) throw new Error('Missing dialog');
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
	const consent = abandoned.querySelector('section[aria-label="Review add-in permissions"]');
	expect(consent?.textContent).toContain('This add-in requests:');
	expect(consent?.textContent).toContain('Read this document');
	expect(consent?.textContent).toContain('https://example.test');
	expect(consent?.textContent).toContain('Lab Notes');
	expect(consent?.textContent).not.toContain('Document.Read.User');
	expect(abandoned.querySelector('a')?.parentElement?.hidden).toBe(true);
	expect([...abandoned.querySelectorAll('p')].map((item) => item.textContent)).toContain(
		`${descriptionPrefix}…`,
	);
	expect(abandoned.querySelector('img, iframe')).toBeNull();
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
	expect((await open('myAddIns')).textContent).toContain('No add-ins');
	activeDialog(editor)?.close();
});
