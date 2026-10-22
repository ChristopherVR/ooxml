// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxEditorElement, registerDocxEditor, type WordCustomization } from './index';
import { resolveWordCustomization } from './customization';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());

function mount(customization?: WordCustomization) {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	if (customization) editor.customization = customization;
	document.body.append(editor);
	return editor;
}
const root = (editor: DocxEditorElement) => editor.shadowRoot!;
const hidden = (editor: DocxEditorElement, label: string) =>
	root(editor).querySelector(`[aria-label="${label}"]`)!.hasAttribute('data-dve-hidden');
type Backstage = HTMLElement & {
	items: readonly { id: string; label: string; group?: string }[];
};
const backstage = (editor: DocxEditorElement) =>
	root(editor).querySelector('office-ui-backstage') as Backstage;
const navIds = (editor: DocxEditorElement) => backstage(editor).items.map((item) => item.id);
const select = (editor: DocxEditorElement, id: string) =>
	backstage(editor).dispatchEvent(
		new CustomEvent('office-backstage-select', { detail: { id }, cancelable: true }),
	);

describe('<docx-editor> customization', () => {
	it('resolves capabilities into ribbon, file, page, shortcut and menu removals', () => {
		const resolved = resolveWordCustomization({
			capabilities: { print: false, copy: false, open: false, settings: false },
		});
		expect(resolved.hiddenActions).toEqual(['print', 'cut', 'copy']);
		expect([...resolved.fileCommands]).toEqual(['print', 'new', 'open']);
		expect([...resolved.backstagePages]).toEqual(['print', 'new', 'open', 'options']);
		expect([...resolved.shortcuts]).toEqual(['print']);
		expect([...resolved.menuItems]).toEqual(['cut', 'copy']);
		expect(resolved.blockCopy).toBe(true);
		expect(resolved.canEdit).toBe(true);
	});

	it('capabilities hide ribbon controls without touching hiddenActions', () => {
		const editor = mount({ capabilities: { print: false, find: false } });
		expect(hidden(editor, 'Print')).toBe(true);
		expect(hidden(editor, 'Find and replace')).toBe(true);
		expect(hidden(editor, 'Bold')).toBe(false);
		expect(editor.hiddenActions).toEqual([]);
		editor.hiddenActions = ['bold'];
		expect(hidden(editor, 'Print')).toBe(true);
		editor.customization = null;
		expect(hidden(editor, 'Print')).toBe(false);
		expect(hidden(editor, 'Bold')).toBe(true);
	});

	it('edit: false makes the document read-only until the capability is removed', () => {
		const editor = mount();
		editor.customization = { capabilities: { edit: false } };
		expect(editor.readOnly).toBe(true);
		editor.readOnly = false;
		expect(editor.readOnly).toBe(true);
		expect(editor.hasAttribute('read-only')).toBe(false);
		editor.customization = {};
		expect(editor.readOnly).toBe(false);
		editor.readOnly = true;
		editor.customization = null;
		expect(editor.readOnly).toBe(true);
	});

	it('removes the File tab pages, Quick Access Save and Share of disabled capabilities', () => {
		const editor = mount({
			capabilities: { open: false, save: false, export: false, print: false, settings: false },
		});
		expect(navIds(editor)).toEqual(['home', 'info']);
		const bar = root(editor).querySelector('office-ui-title-bar') as HTMLElement & {
			state: { quickAccess: { items: { id: string }[] } };
		};
		expect(bar.state.quickAccess.items.map((item) => item.id)).toEqual(['undo', 'redo']);
		editor.customization = { capabilities: { share: false } };
		const actions = root(editor).querySelector('office-ui-ribbon-actions') as HTMLElement & {
			noShare: boolean;
		};
		expect(actions.noShare).toBe(true);
		expect(navIds(editor)).toContain('print');
	});

	it('refuses file commands and shortcuts of disabled capabilities', () => {
		const editor = mount({ capabilities: { print: false, save: false } });
		const commands = vi.fn((event: Event) => event.preventDefault());
		editor.addEventListener('file-command', commands);
		const press = (key: string) => {
			const event = new KeyboardEvent('keydown', {
				key,
				ctrlKey: true,
				bubbles: true,
				cancelable: true,
			});
			editor.dispatchEvent(event);
			return event.defaultPrevented;
		};
		expect(press('p')).toBe(true);
		expect(press('s')).toBe(true);
		expect(commands).not.toHaveBeenCalled();
		editor.customization = null;
		press('s');
		expect(commands).toHaveBeenCalledTimes(1);
	});

	it('blocks copy and cut when copy is off', () => {
		const editor = mount({ capabilities: { copy: false } });
		const fire = (type: string) => {
			const event = new Event(type, { bubbles: true, cancelable: true, composed: true });
			root(editor).querySelector('.dve-canvas')!.dispatchEvent(event);
			return event.defaultPrevented;
		};
		expect(fire('copy')).toBe(true);
		expect(fire('cut')).toBe(true);
		editor.customization = null;
		expect(fire('copy')).toBe(false);
	});

	it('adds host File tab pages, mounts them when shown and tears them down', () => {
		const destroy = vi.fn();
		const mountPage = vi.fn((container: HTMLElement, context: { fileName: string }) => {
			container.textContent = `Policies for ${context.fileName}`;
			return destroy;
		});
		const editor = mount({
			backstage: {
				extraPages: [
					{ id: 'about', label: 'About Acme', group: 'footer', render: 'acme-about' },
					{ id: 'policies', label: 'Policies', render: mountPage },
				],
			},
		});
		expect(navIds(editor)).toEqual([
			'home',
			'info',
			'new',
			'open',
			'save',
			'saveAs',
			'print',
			'export',
			'host.policies',
			'options',
			'host.about',
		]);
		expect(backstage(editor).items.at(-1)).toEqual({
			id: 'host.about',
			label: 'About Acme',
			group: 'footer',
		});
		select(editor, 'host.policies');
		const page = backstage(editor).querySelector<HTMLElement>(
			'[data-backstage-page="host.policies"]',
		)!;
		expect(page.textContent).toBe('Policies for Document1.docx');
		expect(page.getAttribute('data-office-region')).toBe('backstage.host.policies');
		expect(mountPage).toHaveBeenCalledTimes(1);
		select(editor, 'host.about');
		expect(backstage(editor).querySelector('acme-about')).not.toBeNull();
		editor.customization = null;
		expect(destroy).toHaveBeenCalledTimes(1);
		expect(backstage(editor).querySelector('[data-backstage-page^="host."]')).toBeNull();
		expect(navIds(editor)).not.toContain('host.policies');
	});
});
