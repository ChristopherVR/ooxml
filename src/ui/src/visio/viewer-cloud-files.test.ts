// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import type {
	CloudFileProvider,
	CloudOpenedFile,
	DocumentSaveMeta,
	DocumentStorage,
} from 'ooxml-core/embedding';
import { mountViewer } from './binding';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

const provider: CloudFileProvider = {
	id: 'cloud',
	label: 'Cloud',
	account: () => ({ name: 'Ada' }),
	signIn: async () => ({ name: 'Ada' }),
	list: async () => ({ entries: [] }),
	open: async () => {
		throw new Error('unused');
	},
};

const rectangle = (shapeId: string) => ({
	type: 'create-rectangle' as const,
	pageId: '1',
	shapeId,
	x: 2,
	y: 2,
	width: 1,
	height: 1,
});

async function setup(options: Parameters<typeof mountViewer>[1] = {}) {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { fileProviders: [provider], ...options });
	vi.spyOn(viewer.controller, 'parser').mockImplementation(parseVsdx);
	type Edit = (bytes: Uint8Array, commands: Parameters<typeof editVsdx>[1]) => Promise<unknown>;
	vi.spyOn(viewer.controller as unknown as { editor: Edit }, 'editor').mockImplementation(
		async (bytes, commands) => {
			const result = await editVsdx(bytes, commands);
			return { ...result, document: await parseVsdx(result.bytes) };
		},
	);
	return { viewer, root: viewer.element.shadowRoot! };
}

async function cloudFile(
	overrides: Partial<CloudOpenedFile> & { storage: DocumentStorage },
): Promise<CloudOpenedFile> {
	return {
		entry: { id: 'd!1', name: 'Plan.vsdx', type: 'file', kind: 'vsdx' },
		kind: 'vsdx',
		fileName: 'Plan.vsdx',
		bytes: new Uint8Array(await createVsdxFixture()),
		canEdit: true,
		...overrides,
	};
}

const switcher = (root: ShadowRoot) =>
	root.querySelector('office-ui-open-places') as HTMLElement & { shadowRoot: ShadowRoot };
const tab = (root: ShadowRoot, place: string) =>
	switcher(root).shadowRoot.querySelector<HTMLButtonElement>(
		`[role="tab"][data-place="${place}"]`,
	)!;

function choose(root: ShadowRoot, file: CloudOpenedFile): void {
	tab(root, 'cloud').click();
	switcher(root)
		.shadowRoot.querySelector('office-ui-cloud-files')!
		.dispatchEvent(
			new CustomEvent('office-cloud-file-open', {
				detail: { providerId: 'cloud', file },
				bubbles: true,
				composed: true,
			}),
		);
}

const backstageOpen = (root: ShadowRoot) =>
	(root.querySelector('office-ui-backstage') as HTMLElement & { open: boolean }).open;

describe('File > Open with cloud file providers', () => {
	it('shows This device and one place per provider, and restores the plain page', async () => {
		const { viewer, root } = await setup();
		viewer.element.openBackstage('open');
		const places = () =>
			[...(switcher(root).shadowRoot.querySelectorAll('[role="tab"]') ?? [])].map(
				(b) => b.textContent,
			);
		expect(places()).toEqual(['This device', 'Cloud']);
		const browse = root.querySelector<HTMLElement>(
			'office-ui-open-places [data-backstage-action="open"]',
		)!;
		const browseHidden = () => switcher(root).shadowRoot.querySelector('slot')!.hidden;
		expect(browse).not.toBeNull();
		expect(browseHidden()).toBe(false);
		tab(root, 'cloud').click();
		const browser = switcher(root).shadowRoot.querySelector<
			HTMLElement & { accept: string[]; providers: unknown[] }
		>('office-ui-cloud-files')!;
		expect(browser.accept).toEqual(['vsdx']);
		expect(browser.providers).toEqual([provider]);
		expect(browseHidden()).toBe(true);
		tab(root, 'device').click();
		expect(browseHidden()).toBe(false);
		expect(switcher(root).shadowRoot.querySelector('office-ui-cloud-files')).toBeNull();
		viewer.update({ fileProviders: [] });
		expect(places()).toEqual([]);
		viewer.destroy();
	});

	it('translates the place names and the browser labels', async () => {
		const { viewer, root } = await setup();
		viewer.element.setAttribute('lang', 'de');
		viewer.element.openBackstage('open');
		expect(tab(root, 'device').textContent).toBe('Dieses Gerät');
		tab(root, 'cloud').click();
		const browser = switcher(root).shadowRoot.querySelector<
			HTMLElement & { labels: { signIn: string } }
		>('office-ui-cloud-files')!;
		expect(browser.labels.signIn).toBe('Anmelden');
		viewer.destroy();
	});

	it('opens the chosen drawing, saves to its own storage and marks it clean after', async () => {
		const hostSaves: string[] = [];
		const cloudSaves: DocumentSaveMeta[] = [];
		let finish!: () => void;
		const storage: DocumentStorage = {
			save: (_bytes, meta) =>
				new Promise<void>((resolve) => {
					cloudSaves.push(meta);
					finish = resolve;
				}),
		};
		const { viewer, root } = await setup({
			storage: { save: async (_b, meta) => void hostSaves.push(meta.fileName) },
		});
		const commands: string[] = [];
		viewer.element.addEventListener('file-command', (event) =>
			commands.push((event as CustomEvent<{ command: string }>).detail.command),
		);
		viewer.element.openBackstage('open');
		choose(root, await cloudFile({ storage }));
		await vi.waitFor(() => expect(viewer.element.fileName).toBe('Plan.vsdx'));
		expect(commands).toEqual(['open']);
		expect(viewer.element.canEdit).toBe(true);
		await vi.waitFor(() => expect(backstageOpen(root)).toBe(false));

		await viewer.applyEdits([rectangle('20')]);
		expect(viewer.controller.state.edit.dirty).toBe(true);
		viewer.element.openBackstage('save-as');
		root.querySelector<HTMLButtonElement>('[data-backstage-action="download"]')!.click();
		await vi.waitFor(() => expect(cloudSaves).toHaveLength(1));
		expect(viewer.controller.state.edit.dirty).toBe(true);
		finish();
		await vi.waitFor(() => expect(viewer.controller.state.edit.dirty).toBe(false));
		expect(hostSaves).toEqual([]);
		viewer.destroy();
	});

	it('keeps the drawing unsaved when the upload fails', async () => {
		const storage: DocumentStorage = {
			save: async () => {
				throw new Error('Upload failed');
			},
		};
		const { viewer, root } = await setup();
		const errors: Error[] = [];
		viewer.element.addEventListener('document-error', (event) =>
			errors.push((event as CustomEvent<Error>).detail),
		);
		viewer.element.openBackstage('open');
		choose(root, await cloudFile({ storage }));
		await vi.waitFor(() => expect(viewer.element.fileName).toBe('Plan.vsdx'));
		await viewer.applyEdits([rectangle('20')]);
		viewer.element.openBackstage('save-as');
		root.querySelector<HTMLButtonElement>('[data-backstage-action="download"]')!.click();
		await vi.waitFor(() => expect(errors.map((e) => e.message)).toEqual(['Upload failed']));
		expect(viewer.controller.state.edit.dirty).toBe(true);
		viewer.destroy();
	});

	it('opens read-only when the file is, and the next load restores host storage and editing', async () => {
		const cloudSave = vi.fn(async () => undefined);
		const hostSaves: string[] = [];
		const { viewer, root } = await setup({
			storage: { save: async (_b, meta) => void hostSaves.push(meta.fileName) },
		});
		viewer.element.openBackstage('open');
		choose(root, await cloudFile({ canEdit: false, storage: { save: cloudSave } }));
		await vi.waitFor(() => expect(viewer.element.canEdit).toBe(false));
		await expect(viewer.applyEdits([rectangle('20')])).rejects.toThrow('read-only');

		await viewer.load(new File([new Uint8Array(await createVsdxFixture())], 'Other.vsdx'));
		expect(viewer.element.canEdit).toBe(true);
		await viewer.applyEdits([rectangle('21')]);
		viewer.element.openBackstage('save-as');
		root.querySelector<HTMLButtonElement>('[data-backstage-action="download"]')!.click();
		await vi.waitFor(() => expect(hostSaves).toEqual(['Other.vsdx']));
		expect(cloudSave).not.toHaveBeenCalled();
		viewer.destroy();
	});

	it('keeps a host read-only viewer read-only for an editable file', async () => {
		const { viewer, root } = await setup({ readOnly: true });
		viewer.element.openBackstage('open');
		choose(root, await cloudFile({ storage: { save: async () => undefined } }));
		await vi.waitFor(() => expect(viewer.element.fileName).toBe('Plan.vsdx'));
		expect(viewer.element.canEdit).toBe(false);
		viewer.destroy();
	});

	it('does not open when a file-command listener cancels the open', async () => {
		const { viewer, root } = await setup();
		viewer.element.addEventListener('file-command', (event) => event.preventDefault());
		viewer.element.openBackstage('open');
		choose(root, await cloudFile({ storage: { save: async () => undefined } }));
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(viewer.element.fileName).toBe('');
		viewer.destroy();
	});
});
