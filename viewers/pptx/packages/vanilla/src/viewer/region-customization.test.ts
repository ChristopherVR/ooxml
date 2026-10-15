import type { ViewerRegionContext } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createPptxViewer } from './PptxViewer';
import type { Store, ViewerState } from './state';
import type { PptxViewerInstance, PptxViewerOptions } from './types';

let active: PptxViewerInstance[] = [];

function mount(options: PptxViewerOptions = {}): {
	container: HTMLElement;
	viewer: PptxViewerInstance;
} {
	const container = document.createElement('div');
	document.body.appendChild(container);
	const viewer = createPptxViewer(container, options);
	active.push(viewer);
	return { container, viewer };
}

afterEach(() => {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	document.body.replaceChildren();
});

/** The viewer state store (private on the class; used to drive selection). */
function storeOf(viewer: PptxViewerInstance): Store<ViewerState> {
	return (viewer as unknown as { store: Store<ViewerState> }).store;
}

/** The dialog openers the class has but the public instance type leaves out. */
function internals(viewer: PptxViewerInstance): {
	openSettings(): void;
	openShare(): void;
	openPrintDialog(): void;
	openSelectionPane(): void;
	openAccessibility(): void;
	openPasswordProtection(): void;
	openDigitalSignatures(): void;
} {
	return viewer as unknown as ReturnType<typeof internals>;
}

function region(container: ParentNode, id: string): HTMLElement | null {
	return container.querySelector<HTMLElement>(`[data-pptx-region="${id}"]`);
}

describe('vanilla replaced regions: inspector', () => {
	it('mounts, updates on selection change and destroys with the viewer', () => {
		const update = vi.fn();
		const destroy = vi.fn();
		const mountRegion = vi.fn((container: HTMLElement, _context: ViewerRegionContext) => {
			container.textContent = 'host inspector';
			return { update, destroy };
		});
		const { container, viewer } = mount({
			editable: true,
			customization: { regions: { inspector: mountRegion } },
		});
		const slot = region(container, 'inspector');
		expect(slot?.textContent).toBe('host inspector');
		expect(slot?.closest('[data-pptx-inspector]')).not.toBeNull();
		const first = mountRegion.mock.calls[0]?.[1] as ViewerRegionContext;
		expect(first.region).toBe('inspector');
		// The chrome is built before the editable flag lands; the update delivers it.
		expect((update.mock.calls.at(-1)?.[0] as ViewerRegionContext).canEdit).toBe(true);
		expect(first.selectedElementIds).toStrictEqual([]);
		storeOf(viewer).set({ selectedElementIds: ['a'], selectedElementId: 'a' });
		expect(update).toHaveBeenCalled();
		const latest = update.mock.calls.at(-1)?.[0] as ViewerRegionContext;
		expect(latest.selectedElementIds).toStrictEqual(['a']);
		viewer.destroy();
		expect(destroy).toHaveBeenCalledOnce();
	});

	it('keeps a live host (and its container) across an unrelated chrome rebuild', () => {
		const destroy = vi.fn();
		const mountRegion = vi.fn((container: HTMLElement) => {
			container.textContent = 'host';
			return { destroy };
		});
		const { container, viewer } = mount({
			editable: true,
			customization: { regions: { inspector: mountRegion } },
		});
		const before = region(container, 'inspector');
		viewer.hideRibbonTab('draw');
		const after = region(container, 'inspector');
		expect(mountRegion).toHaveBeenCalledOnce();
		expect(destroy).not.toHaveBeenCalled();
		expect(after).toBe(before);
		expect(after?.isConnected).toBe(true);
		expect(after?.closest('[data-pptx-inspector]')).not.toBeNull();
		// A changed renderer destroys the old host; clearing the region destroys the new one.
		const next = vi.fn(() => ({ destroy }));
		viewer.setRegion('inspector', next);
		expect(destroy).toHaveBeenCalledOnce();
		expect(next).toHaveBeenCalledOnce();
		viewer.setRegion('inspector', undefined);
		expect(destroy).toHaveBeenCalledTimes(2);
	});

	it('restores the built-in pane when the region is cleared', () => {
		const { container, viewer } = mount({
			editable: true,
			customization: {
				regions: { inspector: (el: HTMLElement) => void (el.textContent = 'host') },
			},
		});
		expect(region(container, 'inspector')).not.toBeNull();
		viewer.setRegion('inspector', undefined);
		expect(region(container, 'inspector')).toBeNull();
	});
});

describe('vanilla replaced regions: dialogs', () => {
	it('renders dialog.options as a custom element with a settable context', () => {
		class OptionsView extends HTMLElement {
			context: ViewerRegionContext | undefined;
		}
		customElements.define('test-options-view', OptionsView);
		const { viewer } = mount({
			editable: true,
			customization: { regions: { 'dialog.options': 'test-options-view' } },
		});
		internals(viewer).openSettings();
		const element = document.querySelector<OptionsView>('test-options-view');
		expect(element).not.toBeNull();
		expect(element?.context?.region).toBe('dialog.options');
		expect(typeof (element?.context as { setOption?: unknown }).setOption).toBe('function');
		// The built-in rail and panes are not built; the frame is.
		expect(document.querySelector('.pptxv-options-nav')).toBeNull();
		expect(document.querySelector('.pptxv-parity-dialog')).not.toBeNull();
		// Escape dismisses the frame and releases the renderer.
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(document.querySelector('test-options-view')).toBeNull();
	});

	it('routes setOption through the options store and respects locks', () => {
		let context: ViewerRegionContext | undefined;
		const { viewer } = mount({
			editable: true,
			customization: {
				regions: {
					'dialog.options': (_el: HTMLElement, ctx: ViewerRegionContext) => {
						context = ctx;
					},
				},
			},
		});
		viewer.lockSetting('general.userInitials', 'ZZ');
		internals(viewer).openSettings();
		expect(context).toBeDefined();
		const settings = context as unknown as {
			setOption(id: string, value: unknown): void;
			options: { general: { userInitials: string } };
		};
		settings.setOption('general.userInitials', 'AB');
		expect(settings.options.general.userInitials).toBe('ZZ');
	});

	it('destroys an open dialog renderer on viewer destroy', () => {
		const destroy = vi.fn();
		const { viewer } = mount({
			editable: true,
			customization: { regions: { 'dialog.print': () => destroy } },
		});
		internals(viewer).openPrintDialog();
		expect(document.querySelector('[data-pptx-region="dialog.print"]')).not.toBeNull();
		viewer.destroy();
		expect(destroy).toHaveBeenCalledOnce();
	});
});

describe('vanilla replaced regions: other slots', () => {
	it('fills the slides pane, notes and status bar, and keeps the share frame', () => {
		const destroyed: string[] = [];
		const slot =
			(name: string) =>
			(container: HTMLElement): (() => void) => {
				container.textContent = name;
				return () => destroyed.push(name);
			};
		const { container, viewer } = mount({
			editable: true,
			customization: {
				regions: {
					slidesPane: slot('slides'),
					notes: slot('notes'),
					statusBar: slot('status'),
					'dialog.share': slot('share'),
				},
			},
		});
		expect(region(container, 'slidesPane')?.textContent).toBe('slides');
		expect(region(container, 'notes')?.textContent).toBe('notes');
		expect(region(container, 'statusBar')?.textContent).toBe('status');
		expect(region(container, 'statusBar')?.classList.contains('pptxv-statusbar')).toBe(true);
		// Collapsed notes hide the slot (CSS keys on data-collapsed) but keep the host mounted.
		const notesSlot = region(container, 'notes')!.closest<HTMLElement>('.pptxv-notes')!;
		expect(notesSlot.classList.contains('pptxv-region-replaced')).toBe(true);
		expect(notesSlot.dataset.collapsed).toBe('true');
		storeOf(viewer).set({ notesExpanded: true });
		expect(notesSlot.dataset.collapsed).toBe('false');
		expect(region(container, 'notes')?.isConnected).toBe(true);
		internals(viewer).openShare();
		expect(region(document.body, 'dialog.share')?.textContent).toBe('share');
		expect(document.querySelector('.pptxv-modal-panel')).not.toBeNull();
		viewer.destroy();
		expect(destroyed.sort()).toStrictEqual(['notes', 'share', 'slides', 'status']);
	});

	it('does not render a replaced region the host also hides', () => {
		const render = vi.fn();
		const { container } = mount({
			editable: true,
			customization: { hiddenPanels: ['statusBar'], regions: { statusBar: render } },
		});
		expect(render).not.toHaveBeenCalled();
		expect(region(container, 'statusBar')).toBeNull();
	});
});

describe('vanilla host File tab pages', () => {
	it('lists the host label and mounts the renderer when selected', () => {
		const destroy = vi.fn();
		const mountPage = vi.fn((container: HTMLElement, ctx: ViewerRegionContext) => {
			container.textContent = ctx.region;
			return destroy;
		});
		const { container } = mount({
			editable: true,
			customization: {
				backstage: { extraPages: [{ id: 'audit', label: 'Audit log', render: mountPage }] },
			},
		});
		const fileTab = Array.from(
			container.querySelectorAll<HTMLButtonElement>('.pptxv-ribbon-tab'),
		).find((tab) => (tab.textContent ?? '').trim() === 'File');
		fileTab!.click();
		const item = Array.from(
			container.querySelectorAll<HTMLButtonElement>('[data-pptx-backstage-nav-item]'),
		).find((entry) => (entry.textContent ?? '').includes('Audit log'));
		expect(item).toBeDefined();
		// Host pages carry no nav icon.
		expect(item!.querySelector('svg')).toBeNull();
		expect(mountPage).not.toHaveBeenCalled();
		item!.click();
		expect(mountPage).toHaveBeenCalledOnce();
		expect(region(container, 'backstage.host.audit')?.textContent).toBe('backstage.host.audit');
		// Selecting a built-in page releases the host renderer.
		container.querySelector<HTMLButtonElement>('[data-pptx-backstage-nav-item]')!.click();
		expect(destroy).toHaveBeenCalledOnce();
	});
});

describe('vanilla capabilities.edit', () => {
	it('makes the viewer read-only live and restores it', () => {
		const { container, viewer } = mount({ editable: true });
		const root = (): Element | null => container.querySelector('.pptxv');
		expect(root()?.classList.contains('pptxv-editable')).toBe(true);
		viewer.setCapability('edit', false);
		expect(root()?.classList.contains('pptxv-editable')).toBe(false);
		expect(viewer.getMode()).toBe('preview');
		viewer.setCapability('edit', undefined);
		expect(root()?.classList.contains('pptxv-editable')).toBe(true);
		expect(viewer.getMode()).toBe('edit');
	});

	it('starts read-only when the option sets the capability off', () => {
		const { viewer } = mount({ editable: true, customization: { capabilities: { edit: false } } });
		expect(viewer.getMode()).toBe('preview');
	});

	it('ANDs the legacy readOnly flag', () => {
		const { viewer } = mount({ editable: true, readOnly: true });
		expect(viewer.getMode()).toBe('preview');
	});
});

describe('vanilla new hide ids', () => {
	it('does not open a hidden panel or dialog', () => {
		const { container, viewer } = mount({
			editable: true,
			customization: {
				hiddenPanels: ['selectionPane', 'accessibilityPane'],
				hiddenDialogs: ['passwordProtection', 'signatures'],
			},
		});
		internals(viewer).openSelectionPane();
		internals(viewer).openAccessibility();
		internals(viewer).openPasswordProtection();
		internals(viewer).openDigitalSignatures();
		expect(container.querySelector('.pptxv-selection-pane')).toBeNull();
		expect(document.querySelector('.pptxv-parity-dialog')).toBeNull();
	});

	it('carries the ribbon control ids the entry points name', () => {
		const { container } = mount({ editable: true });
		for (const id of [
			'insert.symbols.equation',
			'home.slides.slideTemplates',
			'home.editing.find',
			'review.compare.compare',
			'view.show.selectionPane',
		]) {
			expect(container.querySelector(`[data-ribbon-control="${id}"]`), id).not.toBeNull();
		}
	});
});
