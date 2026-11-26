import { afterEach, expect, it, vi } from 'vitest';
import { openOfficeAddIns } from './add-in-dialog';
import './add-in-local-dialog';

const xml = `<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
	<Id>083a367e-842b-4d4f-ab91-cd5de849e495</Id><Version>1.0.0</Version>
	<ProviderName>Lab Notes</ProviderName><DefaultLocale>en</DefaultLocale>
	<DisplayName DefaultValue="&lt;img src='https://example.test/untrusted' /&gt;"/>
	<Description DefaultValue="Review notes."/>
	<DefaultSettings><SourceLocation DefaultValue="https://example.test/pane"/></DefaultSettings>
	<Permissions>ReadDocument</Permissions>
</OfficeApp>`;

afterEach(() => document.body.replaceChildren());

it('requires consent, keeps the prior install through a failed update and isolates editors', async () => {
	const editor = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const other = document.body
		.appendChild(document.createElement('div'))
		.attachShadow({ mode: 'open' });
	const open = (mode: 'getAddIns' | 'myAddIns', root = editor) => {
		openOfficeAddIns(root, mode, (text) => text, 'word', 'en');
		const dialog = root.querySelector('office-ui-dialog[open]');
		if (!dialog) throw new Error('Missing dialog');
		return dialog;
	};
	const picker = async (root: ReturnType<typeof open>, source: string) => {
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

	const abandoned = open('getAddIns');
	await picker(abandoned, xml);
	await vi.waitFor(() => expect(control(abandoned, 'Install')).toBeDefined());
	expect(abandoned.textContent).toContain('Document.Read.User');
	expect(abandoned.textContent).toContain('https://example.test');
	expect(abandoned.querySelector('img, iframe')).toBeNull();
	abandoned.close();
	expect(open('myAddIns').textContent).toContain('No add-ins');
	editor.querySelector('office-ui-dialog[open]')?.close();

	const get = open('getAddIns');
	await picker(get, xml);
	const install = await vi.waitFor(() => {
		const button = control(get, 'Install');
		if (!button) throw new Error('Missing consent control');
		return button;
	});
	install.click();

	const invalid = open('getAddIns');
	await picker(invalid, xml.replace('https://example.test/pane', 'javascript:alert(1)'));
	await vi.waitFor(() => expect(invalid.textContent).toContain('unsafe URL or origin'));
	expect(control(invalid, 'Update')).toBeUndefined();
	invalid.close();
	const mine = open('myAddIns');
	await vi.waitFor(() => expect(mine.textContent).toContain('1.0.0'));
	expect(mine.textContent).toContain('Running add-ins is unavailable');
	expect(open('myAddIns', other).textContent).toContain('No add-ins');
	other.querySelector('office-ui-dialog[open]')?.close();
	mine.close();

	const newer = open('getAddIns');
	await picker(newer, xml.replace('1.0.0', '1.0.1'));
	await vi.waitFor(() => expect(control(newer, 'Update')).toBeDefined());
	expect(newer.textContent).toContain('Installed version1.0.0');
	control(newer, 'Update')?.click();

	const updated = open('myAddIns');
	await vi.waitFor(() => expect(updated.textContent).toContain('1.0.1'));
	expect(updated.textContent).not.toContain('1.0.0');
	control(updated, 'Remove')?.click();
	expect(open('myAddIns').textContent).toContain('No add-ins');
	editor.querySelector('office-ui-dialog[open]')?.close();
});
