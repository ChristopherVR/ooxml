// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { saveWorkbook } from 'ooxml-core/xlsx/load';
import type {
	CloudFileProvider,
	CloudOpenedFile,
	DocumentSaveMeta,
	DocumentStorage,
} from 'ooxml-core/embedding';
import { newInto, openCloudFile, runFileCommand } from './editor-files';
import { createBackstage, type BackstageHost, type EditorOptions } from './backstage';
import { shellFixture } from './test-support/shell';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.innerHTML = '';
});

const chrome = () => ({ setSaveState: vi.fn(), fileNameChanged: vi.fn() });

async function cloudFile(overrides: Partial<CloudOpenedFile> = {}) {
	const saves: Array<[Uint8Array, DocumentSaveMeta]> = [];
	const storage: DocumentStorage = {
		save: async (bytes, meta) => {
			saves.push([bytes, meta]);
			return undefined;
		},
	};
	const bytes = await saveWorkbook(createWorkbook({ sheets: ['Cloud'] }), 'xlsx');
	const file: CloudOpenedFile = {
		entry: { id: 'd!1', name: 'Cloud.xlsx', type: 'file', kind: 'xlsx' },
		kind: 'xlsx',
		fileName: 'Cloud.xlsx',
		bytes,
		canEdit: true,
		storage,
		...overrides,
	};
	return { file, saves, storage };
}

describe('File > Open from a cloud place', () => {
	it('loads the workbook, names it and saves through the file storage, then marks it clean', async () => {
		const { core } = shellFixture();
		const hostStorage = { save: vi.fn(async () => undefined) };
		core.storage = hostStorage;
		const { file, saves } = await cloudFile();
		const ui = chrome();
		expect(await openCloudFile(core, file, ui)).toBe(true);
		expect(core.fileName).toBe('Cloud.xlsx');
		expect(core.workbook?.sheets[0]?.name).toBe('Cloud');
		core.dirty.set(true);
		await runFileCommand(core, 'save', undefined, { chrome: ui });
		expect(saves).toHaveLength(1);
		expect(saves[0]![1]).toMatchObject({ fileName: 'Cloud.xlsx', reason: 'save' });
		expect(hostStorage.save).not.toHaveBeenCalled();
		expect(core.dirty.dirty).toBe(false);
		await runFileCommand(core, 'saveAs', 'Copy', { chrome: ui });
		expect(saves[1]![1]).toMatchObject({ fileName: 'Copy.xlsx', reason: 'saveAs' });
	});

	it('stays dirty when the upload fails', async () => {
		const { core, events } = shellFixture();
		const { file, storage } = await cloudFile();
		storage.save = () => Promise.reject(new Error('Upload failed'));
		await openCloudFile(core, file, chrome());
		core.dirty.set(true);
		await runFileCommand(core, 'save', undefined, { chrome: chrome() });
		expect(core.dirty.dirty).toBe(true);
		expect(events.filter((event) => event.type === 'workbook-error')).toHaveLength(1);
	});

	it('is read-only for a file that cannot be edited, and a new workbook lifts it', async () => {
		const { core } = shellFixture();
		const { file } = await cloudFile({ canEdit: false });
		await openCloudFile(core, file, chrome());
		expect(core.readOnly).toBe(true);
		expect(core.hostReadOnly).toBe(false);
		newInto(core, chrome());
		expect(core.readOnly).toBe(false);
		expect(core.fileStorage).toBeUndefined();
	});

	it('keeps the host read-only switch when the file is editable', async () => {
		const { core } = shellFixture();
		core.setReadOnly(true);
		const { file } = await cloudFile();
		await openCloudFile(core, file, chrome());
		expect(core.readOnly).toBe(true);
	});

	it('asks before discarding unsaved changes and does nothing when declined', async () => {
		const { core } = shellFixture();
		core.dirty.set(true);
		const { file } = await cloudFile();
		const closed = vi.fn();
		const opened = openCloudFile(core, file, chrome(), closed);
		const dialog = core.host.shadowRoot!.querySelector('office-ui-message-dialog') as HTMLElement;
		expect(dialog.shadowRoot?.querySelector('[role="alertdialog"]')).not.toBeNull();
		(dialog.shadowRoot!.querySelector('button.cancel') as HTMLButtonElement).click();
		expect(await opened).toBe(false);
		expect(closed).not.toHaveBeenCalled();
		expect(core.fileName).not.toBe('Cloud.xlsx');
		expect(core.host.shadowRoot!.querySelector('office-ui-message-dialog')).toBeNull();
	});

	it('opens the file once the discard dialog is confirmed', async () => {
		const { core } = shellFixture();
		core.dirty.set(true);
		const { file } = await cloudFile();
		const closed = vi.fn();
		const opened = openCloudFile(core, file, chrome(), closed);
		const dialog = core.host.shadowRoot!.querySelector('office-ui-message-dialog') as HTMLElement;
		(dialog.shadowRoot!.querySelector('button.confirm') as HTMLButtonElement).click();
		await opened;
		expect(closed).toHaveBeenCalledOnce();
		expect(core.fileName).toBe('Cloud.xlsx');
	});
});

describe('File > Open page places', () => {
	function backstageFor(providers: readonly CloudFileProvider[]) {
		const { core } = shellFixture();
		const opened: CloudOpenedFile[] = [];
		const options: EditorOptions = {
			locale: 'en',
			theme: 'auto',
			author: 'Ada',
			calculation: 'automatic',
			iteration: undefined,
		};
		const host: BackstageHost = {
			ctx: core.ctx,
			fileName: () => 'Book1.xlsx',
			saveState: () => 'saved',
			fileCommand: vi.fn(),
			newFromTemplate: vi.fn(),
			close: vi.fn(),
			options: () => options,
			setOption: vi.fn(),
			setProperty: vi.fn(),
			hiddenActions: () => [],
			setHiddenActions: vi.fn(),
			fileProviders: () => providers,
			openCloudFile: (file) => opened.push(file),
		};
		const backstage = createBackstage(host);
		core.ctx.root.append(backstage.element);
		return { backstage, opened, el: backstage.element };
	}
	const provider = (): CloudFileProvider => ({
		id: 'onedrive',
		label: 'OneDrive',
		account: () => null,
		signIn: async () => ({ name: 'Ada' }),
		list: async () => ({ entries: [] }),
		open: async () => {
			throw new Error('unused');
		},
	});

	it('shows only Browse without providers', () => {
		const { backstage, el } = backstageFor([]);
		backstage.open('open');
		expect(el.querySelector('.xve-backstage-place-tab')).toBeNull();
		expect(el.textContent).toContain('Browse...');
	});

	it('offers This device and one place per provider, and forwards the opened file', async () => {
		const { backstage, el, opened } = backstageFor([provider()]);
		backstage.open('open');
		const tabs = [...el.querySelectorAll<HTMLButtonElement>('.xve-backstage-place-tab')];
		expect(tabs.map((tab) => tab.textContent)).toEqual(['This device', 'OneDrive']);
		expect(el.textContent).toContain('Browse...');
		tabs[1]!.click();
		const cloud = el.querySelector('office-ui-cloud-files') as HTMLElement & {
			accept: string[];
			labels: { signIn: string };
		};
		expect(cloud.accept).toEqual(['xlsx']);
		expect(cloud.labels.signIn).toBe('Sign in');
		const { file } = await cloudFile();
		cloud.dispatchEvent(
			new CustomEvent('office-cloud-file-open', {
				detail: { providerId: 'onedrive', file },
				bubbles: true,
				composed: true,
			}),
		);
		expect(opened).toEqual([file]);
	});
});
