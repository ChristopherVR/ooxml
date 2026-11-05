// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument, saveDocx } from 'ooxml-core/docx';
import type {
	CloudEntry,
	CloudFileProvider,
	CloudOpenedFile,
	DocumentStorage,
} from 'ooxml-core/embedding';
import type { EditorChrome } from './editor-chrome';
import { DocxEditorElement, registerDocxEditor } from './index';
import { cloudFileLabels } from './cloud-labels';
import { strings } from './localization-strings';

registerDocxEditor();

type Updating = HTMLElement & { updateComplete?: Promise<unknown> };
const chrome = (editor: DocxEditorElement) =>
	(editor as unknown as { core: { shell: { chrome: EditorChrome } } }).core.shell.chrome;

function mount(): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = createDocument();
	editor.fileName = 'Report.docx';
	document.body.append(editor);
	return editor;
}

const entry: CloudEntry = { id: 'd!1', name: 'Plan.docx', type: 'file', kind: 'docx' };

function provider(): CloudFileProvider {
	return {
		id: 'fake',
		label: 'Fake Drive',
		account: () => ({ name: 'Ann' }),
		signIn: async () => ({ name: 'Ann' }),
		list: async () => ({ entries: [entry] }),
		open: async () => {
			throw new Error('not used');
		},
	};
}

async function opened(canEdit: boolean, storage: DocumentStorage): Promise<CloudOpenedFile> {
	return {
		entry,
		kind: 'docx',
		fileName: 'Plan.docx',
		bytes: await saveDocx(createDocument()),
		canEdit,
		storage,
	};
}

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

describe('<docx-editor> cloud file places', () => {
	it('keeps the plain Browse page without providers', () => {
		const editor = mount();
		const page = (n: string) =>
			chrome(editor).backstage.element.querySelector(`[data-backstage-page="${n}"]`);
		chrome(editor).backstage.open('open');
		expect(page('open')?.querySelector('.dve-open-places')).toBeNull();
		expect(page('open')?.textContent).toContain('Browse');
	});

	it('lists this device and one place per provider', async () => {
		const editor = mount();
		editor.fileProviders = [provider()];
		chrome(editor).backstage.open('open');
		const page = chrome(editor).backstage.element.querySelector('[data-backstage-page="open"]')!;
		const places = [...page.querySelectorAll('.dve-open-place')].map((b) => b.textContent);
		expect(places).toEqual(['This device', 'Fake Drive']);
		(page.querySelectorAll('.dve-open-place')[1] as HTMLButtonElement).click();
		const files = page.querySelector('office-ui-cloud-files') as Updating & {
			accept: string[];
			labels: Record<string, string>;
		};
		expect(files.accept).toEqual(['docx']);
		expect(files.labels.signIn).toBe('Sign in');
	});

	it('loads an opened file, saves through its own storage and returns to the host storage', async () => {
		const editor = mount();
		const host: DocumentStorage = { save: vi.fn(async () => undefined) };
		editor.storage = host;
		const cloud: DocumentStorage = { save: vi.fn(async () => undefined) };
		chrome(editor).backstage.open('open');
		await chrome(editor).openCloudFile(await opened(true, cloud));
		expect(editor.fileName).toBe('Plan.docx');
		expect(editor.dirty).toBe(false);
		expect(chrome(editor).backstage.isOpen).toBe(false);
		await chrome(editor).run('save');
		expect(cloud.save).toHaveBeenCalledTimes(1);
		expect(host.save).not.toHaveBeenCalled();
		await editor.load(await saveDocx(createDocument()));
		await chrome(editor).run('save');
		expect(host.save).toHaveBeenCalledTimes(1);
		expect(cloud.save).toHaveBeenCalledTimes(1);
	});

	it('stays dirty when the cloud save fails', async () => {
		const editor = mount();
		const errors: unknown[] = [];
		editor.addEventListener('document-error', (event) =>
			errors.push((event as CustomEvent).detail),
		);
		await chrome(editor).openCloudFile(
			await opened(true, { save: () => Promise.reject(new Error('conflict')) }),
		);
		await chrome(editor).run('save');
		expect(errors).toHaveLength(1);
		expect(editor.dirty).toBe(true);
	});

	it('locks a file that cannot be edited, until another document loads', async () => {
		const editor = mount();
		const changes: boolean[] = [];
		editor.addEventListener('readonly-change', (event) =>
			changes.push((event as CustomEvent).detail.readOnly),
		);
		await chrome(editor).openCloudFile(await opened(false, { save: async () => undefined }));
		expect(editor.readOnly).toBe(true);
		expect(changes).toEqual([true]);
		await editor.load(await saveDocx(createDocument()));
		expect(editor.readOnly).toBe(false);
	});

	it('keeps a host readOnly switch when the file can be edited', async () => {
		const editor = mount();
		editor.readOnly = true;
		await chrome(editor).openCloudFile(await opened(true, { save: async () => undefined }));
		expect(editor.readOnly).toBe(true);
	});

	it('does not open when the host removed the Open command', async () => {
		const editor = mount();
		const load = vi.spyOn(editor, 'load');
		editor.customization = { capabilities: { open: false } };
		await chrome(editor).openCloudFile(await opened(true, { save: async () => undefined }));
		expect(load).not.toHaveBeenCalled();
	});
});

describe('cloud file labels', () => {
	it('define every label in every locale', () => {
		for (const locale of Object.keys(strings) as Array<keyof typeof strings>) {
			const labels = cloudFileLabels(locale);
			expect(Object.keys(labels)).toHaveLength(26);
			for (const value of Object.values(labels)) expect(value.trim()).not.toBe('');
		}
		expect(cloudFileLabels('de').signIn).toBe('Anmelden');
	});
});

describe('<docx-editor> New with unsaved changes', () => {
	const dialogIn = (editor: DocxEditorElement) =>
		editor.shadowRoot?.querySelector('office-ui-message-dialog') ?? null;

	it('asks in an in-app dialog and starts the new document only after Discard', async () => {
		const editor = mount();
		chrome(editor).setSaveState('dirty');
		const running = chrome(editor).run('new');
		const dialog = dialogIn(editor) as HTMLElement;
		expect(dialog.shadowRoot?.querySelector('[role="alertdialog"]')).not.toBeNull();
		expect(dialog.shadowRoot?.querySelector('.message')?.textContent).toBe(
			'Discard unsaved changes and start a new document?',
		);
		expect(dialog.shadowRoot?.querySelector('button.confirm')?.textContent).toBe('Discard');
		(dialog.shadowRoot!.querySelector('button.cancel') as HTMLButtonElement).click();
		await running;
		expect(editor.fileName).toBe('Report.docx');
		expect(dialogIn(editor)).toBeNull();

		const again = chrome(editor).run('new');
		(
			(dialogIn(editor) as HTMLElement).shadowRoot!.querySelector(
				'button.confirm',
			) as HTMLButtonElement
		).click();
		await again;
		expect(editor.fileName).not.toBe('Report.docx');
		expect(editor.dirty).toBe(false);
	});
});
