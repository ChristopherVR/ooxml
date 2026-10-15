// @vitest-own-worker: mounts several full viewers; see region-customization.test-support.ts.
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	destroyMounted,
	internals,
	mount,
	region,
	storeOf,
} from './region-customization.test-support';
import type { ViewerRegionContext } from './region-customization.test-support';

afterEach(destroyMounted);

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
