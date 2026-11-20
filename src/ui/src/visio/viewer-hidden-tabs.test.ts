// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import { VISIO_RIBBON_TAB_IDS, demoDocument } from 'ooxml-core/visio/ui';
import type { DocumentSaveMeta, DocumentStorage } from 'ooxml-core/embedding';
import { mountViewer } from './binding';
import type { VisioCustomization } from './index';
import { RIBBON_TABS } from './ribbon';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

const NO_FILE: VisioCustomization = { ribbon: { hiddenTabs: ['file'] } };

function mount(options: Parameters<typeof mountViewer>[1] = {}) {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: structuredClone(demoDocument), ...options });
	return { viewer, root: viewer.element.shadowRoot! };
}
type Ribbon = HTMLElement & {
	noFile: boolean;
	selected: string;
	fileButton(): HTMLButtonElement | null;
	updateComplete: Promise<unknown>;
};
const ribbonOf = (root: ShadowRoot) => root.querySelector('office-ui-ribbon') as Ribbon;
const backstageOf = (root: ShadowRoot) =>
	root.querySelector('office-ui-backstage') as HTMLElement & { open: boolean };
const tabIds = async (root: ShadowRoot) => {
	await new Promise((resolve) => setTimeout(resolve));
	await ribbonOf(root).updateComplete;
	return [...ribbonOf(root).shadowRoot!.querySelectorAll<HTMLElement>('[role="tab"]')].map(
		(tab) => tab.dataset.tab,
	);
};

describe('<visio-viewer> ribbon.hiddenTabs', () => {
	it('names every native tab the ribbon builds', () => {
		expect(['file', ...RIBBON_TABS.map(([id]) => id)]).toEqual([...VISIO_RIBBON_TAB_IDS]);
	});

	it('removes File, its key tip and the backstage while the other tabs stay', async () => {
		const { viewer, root } = mount({ customization: NO_FILE });
		const ribbon = ribbonOf(root);
		await ribbon.updateComplete;
		expect(ribbon.noFile).toBe(true);
		expect(ribbon.fileButton()).toBeNull();
		const file = ribbon.shadowRoot!.querySelector<HTMLElement>('.file')!;
		expect(file.hidden).toBe(true);
		expect(file.hasAttribute('data-keytip')).toBe(false);
		expect(await tabIds(root)).toEqual([
			'home',
			'insert',
			'design',
			'data',
			'process',
			'review',
			'view',
			'help',
		]);
		ribbon.dispatchEvent(new CustomEvent('office-ribbon-file'));
		expect(backstageOf(root).open).toBe(false);
		viewer.element.openBackstage('info');
		expect(backstageOf(root).open).toBe(false);
		viewer.destroy();
	});

	it('keeps Quick Access Save, storage and file-command working', async () => {
		const saves: DocumentSaveMeta[] = [];
		const storage: DocumentStorage = { save: async (_bytes, meta) => void saves.push(meta) };
		const { viewer, root } = mount({ customization: NO_FILE, storage });
		vi.spyOn(viewer.controller, 'parser').mockImplementation(parseVsdx);
		await viewer.createBlankDrawing();
		const commands: string[] = [];
		viewer.element.addEventListener('file-command', (event) =>
			commands.push((event as CustomEvent<{ command: string }>).detail.command),
		);
		root
			.querySelector('office-ui-title-bar')!
			.dispatchEvent(new CustomEvent('office-command', { detail: { command: 'save' } }));
		await vi.waitFor(() => expect(saves).toHaveLength(1));
		expect(commands).toEqual(['save']);
		expect(saves[0]?.reason).toBe('save');
		expect(backstageOf(root).open).toBe(false);
		viewer.destroy();
	});

	it('closes an open backstage when File is hidden at runtime and returns focus to the ribbon', async () => {
		const { viewer, root } = mount();
		const ribbon = ribbonOf(root);
		await ribbon.updateComplete;
		ribbon.fileButton()!.click();
		expect(backstageOf(root).open).toBe(true);
		viewer.update({ customization: NO_FILE });
		expect(backstageOf(root).open).toBe(false);
		expect((ribbon.shadowRoot!.activeElement as HTMLElement | null)?.getAttribute('role')).toBe(
			'tab',
		);
		viewer.update({ customization: null });
		expect(ribbon.noFile).toBe(false);
		ribbon.fileButton()!.click();
		expect(backstageOf(root).open).toBe(true);
		viewer.destroy();
	});

	it('hides other native tabs and selects the first one left', async () => {
		const { viewer, root } = mount({ customization: { ribbon: { hiddenTabs: ['home', 'data'] } } });
		expect(await tabIds(root)).toEqual(['insert', 'design', 'process', 'review', 'view', 'help']);
		expect(ribbonOf(root).selected).toBe('insert');
		viewer.update({ customization: {} });
		expect(await tabIds(root)).toContain('home');
		viewer.destroy();
	});

	it('leaves commands that live only on a hidden tab out of Tell me', () => {
		const { viewer, root } = mount();
		const search = root.querySelector('.tell-me') as HTMLElement & { commands: { id: string }[] };
		const ids = () => {
			search.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
			return search.commands.map((command) => command.id);
		};
		const inProcess = [...root.querySelectorAll('[data-ribbon-tab="process"] [command]')].map(
			(node) => node.getAttribute('command')!,
		);
		const elsewhere = new Set(
			[...root.querySelectorAll('.toolbar [command]')]
				.filter((node) => !node.closest('[data-ribbon-tab="process"]'))
				.map((node) => node.getAttribute('command')),
		);
		const only = inProcess.filter((id) => !elsewhere.has(id) && ids().includes(id));
		expect(only.length).toBeGreaterThan(0);
		viewer.update({ customization: { ribbon: { hiddenTabs: ['process'] } } });
		const hidden = ids();
		for (const id of only) expect(hidden).not.toContain(id);
		expect(hidden).toContain('copy');
		viewer.update({ customization: null });
		for (const id of only) expect(ids()).toContain(id);
		viewer.destroy();
	});
});
