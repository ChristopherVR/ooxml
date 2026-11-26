// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { setOfficeAddInCatalog } from '../ribbon/add-in-dialog';
import type { RibbonAddInTab } from '../ribbon/add-in-tabs';
import type { OfficeUiRibbon } from '../ribbon/ribbon-tabs';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());

const tab = (run: () => void): RibbonAddInTab => ({
	id: 'pdf',
	label: 'PDF',
	groups: [{ label: 'Create PDF', commands: [{ id: 'create-pdf', label: 'Create PDF', run }] }],
});

describe('Word ribbon add-in tabs', () => {
	it('installs and removes a catalog add-in through the editor’s two commands', async () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const pdf = tab(vi.fn());
		const install = vi.fn(() => {
			editor.ribbonAddIns = [pdf];
		});
		const remove = vi.fn(() => {
			editor.ribbonAddIns = [];
		});
		document.body.append(editor);
		const root = editor.shadowRoot!;
		const ribbon = root.querySelector<OfficeUiRibbon>('office-ui-ribbon')!;
		setOfficeAddInCatalog(ribbon, [{ id: pdf.id, label: pdf.label, install, remove }]);
		await vi.waitFor(() => expect(ribbon.tabButton('insert')).toBeTruthy());
		ribbon.tabButton('insert')!.click();
		const dialog = () => root.querySelector('office-ui-add-ins-dialog')?.shadowRoot;
		const click = (label: string) => {
			const button = root.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`)!;
			button.focus();
			button.click();
		};
		const clickDialog = async (label: string) => {
			const button = await vi.waitFor(() => {
				const button = dialog()?.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`);
				if (!button) throw new Error(`Missing ${label}`);
				return button;
			});
			button.click();
		};
		click('Get Add-ins');
		const admin = await vi.waitFor(() => {
			const pivot = dialog()?.querySelector<HTMLButtonElement>('#pivot-1');
			if (!pivot) throw new Error('Missing Admin Managed');
			return pivot;
		});
		expect(admin.textContent).toBe('Admin Managed');
		admin.click();
		await clickDialog('Install PDF');
		await vi.waitFor(() => expect(dialog()).toBeUndefined());
		expect(install).toHaveBeenCalledOnce();
		expect(ribbon.tabButton('pdf')).toBeTruthy();
		click('My Add-ins');
		await clickDialog('Remove PDF');
		await vi.waitFor(() => expect(dialog()).toBeUndefined());
		expect(remove).toHaveBeenCalledOnce();
		expect(ribbon.tabButton('pdf')).toBeNull();
	});

	it('opens a host-provided add-in from Insert > My Add-ins', async () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.ribbonAddIns = [tab(vi.fn())];
		document.body.append(editor);
		const root = editor.shadowRoot!;
		const ribbon = root.querySelector<OfficeUiRibbon>('office-ui-ribbon')!;
		await vi.waitFor(() => expect(ribbon.tabButton('pdf')).toBeTruthy());
		ribbon.tabButton('insert')!.click();
		root.querySelector<HTMLButtonElement>('button[aria-label="My Add-ins"]')!.click();
		const open = await vi.waitFor(() => {
			const button = root
				.querySelector('office-ui-add-ins-dialog')
				?.shadowRoot?.querySelector<HTMLButtonElement>('[aria-label="Open PDF"]');
			if (!button) throw new Error('Missing add-in');
			return button;
		});
		open.click();
		await vi.waitFor(() => expect(ribbon.selected).toBe('pdf'));
	});
	it('keeps tabs set before the editor connects and reports their commands', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const run = vi.fn();
		const heard = vi.fn();
		editor.addEventListener('office-ribbon-add-in', (event) =>
			heard((event as CustomEvent).detail),
		);
		const action = vi.fn();
		editor.addEventListener('ribbon-action', action);
		editor.ribbonAddIns = [tab(run), { id: 'home', label: 'Not Home', groups: [] }];
		document.body.append(editor);
		const ribbon = editor.shadowRoot!.querySelector('office-ui-ribbon')!;
		const panels = [...ribbon.querySelectorAll<HTMLElement>('[data-ribbon-tab]')];
		expect(panels.filter((panel) => panel.dataset.ribbonTab === 'home')).toHaveLength(1);
		const panel = panels.at(-1)!;
		expect(panel.dataset.label).toBe('PDF');
		expect(panel.className).toBe('ribbon-panel');
		panel.querySelector('[data-add-in-command="create-pdf"]')!.dispatchEvent(
			new CustomEvent('office-command', {
				detail: { command: 'create-pdf' },
				bubbles: true,
				composed: true,
			}),
		);
		expect(run).toHaveBeenCalledTimes(1);
		expect(heard).toHaveBeenCalledWith({ tab: 'pdf', command: 'create-pdf' });
		expect(action).not.toHaveBeenCalled();
		// A locale change translates Word's tabs and leaves the host's label alone.
		editor.locale = 'fr';
		expect(panel.dataset.label).toBe('PDF');
		expect(panel.isConnected).toBe(true);
		editor.ribbonAddIns = [];
		expect(ribbon.querySelector('[data-add-in]')).toBeNull();
		expect(editor.ribbonAddIns).toEqual([]);
	});
});
