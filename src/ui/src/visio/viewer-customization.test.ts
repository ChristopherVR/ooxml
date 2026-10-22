import { afterEach, describe, expect, it, vi } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { mountViewer } from './binding';
import type { VisioCustomization } from './index';
import { resolveVisioCustomization, visioActionBlocked } from './viewer-customization';

afterEach(() => document.body.replaceChildren());

function mount(customization?: VisioCustomization) {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, {
		document: structuredClone(demoDocument),
		...(customization ? { customization } : {}),
	});
	return { viewer, root: viewer.element.shadowRoot! };
}
type Backstage = HTMLElement & { items: readonly { id: string; label: string; group?: string }[] };
const backstage = (root: ShadowRoot) => root.querySelector('office-ui-backstage') as Backstage;
const navIds = (root: ShadowRoot) => backstage(root).items.map((item) => item.id);

describe('<visio-viewer> customization', () => {
	it('resolves capabilities and refuses the actions they govern', () => {
		const resolved = resolveVisioCustomization({ capabilities: { copy: false, find: false } });
		expect(visioActionBlocked(resolved, { type: 'clipboard', operation: 'copy' })).toBe(true);
		expect(visioActionBlocked(resolved, { type: 'clipboard', operation: 'paste' })).toBe(false);
		expect(visioActionBlocked(resolved, { type: 'search' })).toBe(true);
		expect(visioActionBlocked(resolved, { type: 'grid' })).toBe(false);
		expect(resolved.commands.has('ctx-copy')).toBe(true);
	});

	it('hides governed commands and shows them again when the capability is removed', () => {
		const { viewer, root } = mount({ capabilities: { copy: false, comment: false } });
		const cut = root.querySelector<HTMLElement>('[command="cut"]')!;
		const comment = root.querySelector<HTMLElement>('[command="ctx-comment"]')!;
		expect(cut.hidden).toBe(true);
		expect(comment.hidden).toBe(true);
		expect(root.querySelector<HTMLElement>('[command="paste"]')?.hidden ?? false).toBe(false);
		viewer.update({ customization: null });
		expect(cut.hidden).toBe(false);
		expect(comment.hidden).toBe(false);
		viewer.destroy();
	});

	it('edit: false overrides canEdit until it is removed', () => {
		const { viewer } = mount({ capabilities: { edit: false } });
		expect(viewer.element.canEdit).toBe(false);
		viewer.update({ canEdit: true });
		expect(viewer.element.canEdit).toBe(false);
		viewer.update({ customization: {} });
		expect(viewer.element.canEdit).toBe(true);
		viewer.destroy();
	});

	it('removes the File tab pages of disabled capabilities', () => {
		const { viewer, root } = mount({
			capabilities: { open: false, save: false, print: false, share: false, settings: false },
		});
		expect(navIds(root)).toEqual(['info', 'export', 'close', 'account', 'feedback']);
		viewer.update({ customization: null });
		expect(navIds(root)).toContain('print');
		viewer.destroy();
	});

	it('adds host File tab pages, mounts them on show and tears them down', () => {
		const destroy = vi.fn();
		const render = vi.fn((container: HTMLElement, context: { pageCount: number }) => {
			container.textContent = `${context.pageCount} pages`;
			return { destroy };
		});
		const { viewer, root } = mount({
			backstage: {
				extraPages: [
					{ id: 'about', label: 'About', group: 'footer', render: 'acme-about' },
					{ id: 'policies', label: 'Policies', render },
				],
			},
		});
		const ids = navIds(root);
		expect(ids.indexOf('host.policies')).toBe(ids.indexOf('close') + 1);
		expect(ids.at(-1)).toBe('host.about');
		backstage(root).dispatchEvent(
			new CustomEvent('office-backstage-select', {
				detail: { id: 'host.policies' },
				cancelable: true,
			}),
		);
		const page = backstage(root).querySelector('[data-backstage-page="host.policies"]')!;
		expect(page.textContent).toBe(`${demoDocument.pages.length} pages`);
		viewer.update({ customization: null });
		expect(destroy).toHaveBeenCalledTimes(1);
		expect(navIds(root)).not.toContain('host.policies');
		viewer.destroy();
	});

	it('adds host commands to the shape and page menus and runs them with the selection', () => {
		const onSelect = vi.fn();
		const { viewer, root } = mount({
			contextMenu: {
				extraCommands: [
					{ id: 'send', label: 'Send to chat', onSelect },
					{ id: 'shape-only', label: 'Shape only', when: (c) => c.target === 'shape', onSelect },
				],
			},
		});
		const shape = root.querySelector<SVGGElement>('svg.paper [data-shape-id="s1"]')!;
		shape.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, clientX: 30, clientY: 40 }));
		const menu = root.querySelector<HTMLElement>('[data-context-menu="shape"]')!;
		const first = menu.firstElementChild!;
		expect(first.getAttribute('command')).toBe('host:send');
		expect(menu.querySelector('[command="host:shape-only"]')).not.toBeNull();
		first.shadowRoot!.querySelector('button')!.click();
		expect(onSelect).toHaveBeenCalledWith(
			expect.objectContaining({
				target: 'shape',
				pageIndex: 0,
				shapes: [expect.objectContaining({ id: 's1' })],
			}),
		);
		root
			.querySelector('.viewport')!
			.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, clientX: 5, clientY: 5 }));
		const page = root.querySelector<HTMLElement>('[data-context-menu="page"]')!;
		expect(page.querySelector('[command="host:send"]')).not.toBeNull();
		expect(page.querySelector('[command="host:shape-only"]')).toBeNull();
		viewer.destroy();
	});
});
