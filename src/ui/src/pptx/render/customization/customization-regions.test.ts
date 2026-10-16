// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { PowerPointViewerAPI } from '../../types';
import { BACKSTAGE_NAV } from '../backstage';
import { VIEWER_REGION_IDS } from './customization-catalog';
import { createCustomizationController } from './customization-controller';
import { VIEWER_REGION_ATTRIBUTE, ViewerRegionHost } from './customization-regions';
import type { ViewerRegionContext } from './customization-regions';
import {
	getRegionRenderer,
	isDialogAvailable,
	isEditingAllowed,
	isRegionReplaced,
	resolveCustomization,
} from './customization-resolve';
import { backstageNavEntries } from './customization-surfaces';

const viewer = {} as PowerPointViewerAPI;

function context(overrides: Partial<ViewerRegionContext> = {}): ViewerRegionContext {
	return {
		region: 'inspector',
		viewer,
		mode: 'edit',
		canEdit: true,
		activeSlideIndex: 0,
		slideCount: 3,
		selectedElementIds: [],
		locale: 'en',
		close: () => {},
		...overrides,
	};
}

describe('ViewerRegionHost', () => {
	afterEach(() => {
		vi.restoreAllMocks();
		document.body.replaceChildren();
	});

	it('mounts a function once, updates on change and tears down', () => {
		const container = document.createElement('div');
		const update = vi.fn();
		const destroy = vi.fn();
		const mount = vi.fn((target: HTMLElement) => {
			target.textContent = 'host';
			return { update, destroy };
		});
		const host = new ViewerRegionHost(container);
		host.render(mount, context());
		host.render(mount, context());
		expect(mount).toHaveBeenCalledTimes(1);
		expect(update).not.toHaveBeenCalled();
		expect(container.getAttribute(VIEWER_REGION_ATTRIBUTE)).toBe('inspector');
		expect(container.textContent).toBe('host');

		host.render(mount, context({ selectedElementIds: ['a'] }));
		expect(update).toHaveBeenCalledTimes(1);
		expect(update.mock.calls[0]?.[0].selectedElementIds).toStrictEqual(['a']);
		host.render(mount, context({ selectedElementIds: ['a'] }));
		expect(update).toHaveBeenCalledTimes(1);

		host.destroy();
		expect(destroy).toHaveBeenCalledTimes(1);
	});

	it('accepts a cleanup function and remounts when the renderer changes', () => {
		const container = document.createElement('div');
		const cleanup = vi.fn();
		const host = new ViewerRegionHost(container);
		host.render(() => cleanup, context());
		const second = vi.fn();
		host.render(second, context());
		expect(cleanup).toHaveBeenCalledTimes(1);
		expect(second).toHaveBeenCalledTimes(1);
	});

	it('gives the renderer callbacks that always reach the latest binding callback', () => {
		const container = document.createElement('div');
		let kept: ViewerRegionContext | undefined;
		const mount = (_: HTMLElement, ctx: ViewerRegionContext) => {
			kept = ctx;
		};
		const first = vi.fn();
		const latest = vi.fn();
		const host = new ViewerRegionHost(container);
		host.render(mount, context({ close: first }));
		host.render(mount, context({ close: latest }));
		kept?.close();
		expect(first).not.toHaveBeenCalled();
		expect(latest).toHaveBeenCalledTimes(1);
	});

	it('creates a custom element from a tag or a spec and keeps its context current', () => {
		const container = document.createElement('div');
		const host = new ViewerRegionHost(container);
		host.render({ element: 'x-settings', properties: { flavour: 'kiosk' } }, context());
		const element = container.firstElementChild as HTMLElement & {
			flavour?: string;
			context?: ViewerRegionContext;
		};
		expect(element.localName).toBe('x-settings');
		expect(element.flavour).toBe('kiosk');
		host.render(
			{ element: 'x-settings', properties: { flavour: 'kiosk' } },
			context({ activeSlideIndex: 2 }),
		);
		// An equal spec in a new object is the same renderer: same element, new context.
		expect(container.firstElementChild).toBe(element);
		expect(element.context?.activeSlideIndex).toBe(2);
		host.render({ element: 'x-settings', properties: { flavour: 'other' } }, context());
		expect(container.firstElementChild).not.toBe(element);

		const tag = 'x-inspector';
		host.render(tag, context());
		host.render(tag, context({ canEdit: false }));
		const inspector = container.firstElementChild as HTMLElement & {
			context?: ViewerRegionContext;
		};
		expect(container.children).toHaveLength(1);
		expect(inspector.localName).toBe(tag);
		expect(inspector.context?.canEdit).toBeFalsy();
		host.destroy();
		expect(container.children).toHaveLength(0);
	});

	it('contains a renderer that throws', () => {
		const error = vi.spyOn(console, 'error').mockImplementation(() => {});
		const host = new ViewerRegionHost(document.createElement('div'));
		host.render(() => {
			throw new Error('boom');
		}, context());
		expect(error).toHaveBeenCalledTimes(1);
		expect(() => host.destroy()).not.toThrow();
	});
});

describe('regions and capabilities in the resolver', () => {
	it('lists every built-in File page as a region', () => {
		for (const item of BACKSTAGE_NAV) {
			expect(VIEWER_REGION_IDS).toContain(`backstage.${item.id}`);
		}
		expect(new Set(VIEWER_REGION_IDS).size).toBe(VIEWER_REGION_IDS.length);
	});

	it('keeps replaced regions unless the same region is hidden', () => {
		const mount = () => {};
		const resolved = resolveCustomization({
			regions: {
				inspector: mount,
				notes: mount,
				'dialog.options': 'my-settings',
				'dialog.print': mount,
				'backstage.account': mount,
			},
			hiddenPanels: ['notes'],
			capabilities: { print: false },
			backstage: { hiddenPages: ['account'] },
		});
		expect(getRegionRenderer(resolved, 'inspector')).toBe(mount);
		expect(getRegionRenderer(resolved, 'dialog.options')).toBe('my-settings');
		expect(isRegionReplaced(resolved, 'notes')).toBeFalsy();
		expect(isRegionReplaced(resolved, 'dialog.print')).toBeFalsy();
		expect(isRegionReplaced(resolved, 'backstage.account')).toBeFalsy();
	});

	it('expands capabilities into every entry point', () => {
		const resolved = resolveCustomization({
			capabilities: {
				edit: false,
				export: false,
				share: false,
				copy: false,
				present: false,
				comment: false,
			},
		});
		expect(isEditingAllowed(resolved)).toBeFalsy();
		expect(isDialogAvailable(resolved, 'export')).toBeFalsy();
		expect(isDialogAvailable(resolved, 'share')).toBeFalsy();
		expect(isDialogAvailable(resolved, 'broadcast')).toBeFalsy();
		expect(isDialogAvailable(resolved, 'print')).toBeTruthy();
		expect(resolved.hiddenActions.has('export')).toBeTruthy();
		expect(resolved.hiddenActions.has('slideShow')).toBeTruthy();
		expect(resolved.hiddenElementCommands.has('copy')).toBeTruthy();
		expect(resolved.hiddenElementCommands.has('comment')).toBeTruthy();
		expect(resolved.keyboard.disabled.has('copy')).toBeTruthy();
		expect(resolved.hiddenRibbonControls.has('home.clipboard.cut')).toBeTruthy();
		expect(resolved.hiddenRibbonGroups.has('review.comments')).toBeTruthy();
		expect(resolved.hiddenBackstageCards.has('copyImage')).toBeTruthy();
		expect(isEditingAllowed(resolveCustomization({ capabilities: { edit: true } }))).toBeTruthy();
	});

	it('hides the entry points of a hidden panel or dialog', () => {
		const resolved = resolveCustomization({
			hiddenPanels: ['findReplace', 'selectionPane'],
			hiddenDialogs: ['hyperlink', 'passwordProtection', 'options'],
		});
		expect(resolved.hiddenRibbonControls.has('home.editing.find')).toBeTruthy();
		expect(resolved.keyboard.disabled.has('findReplace')).toBeTruthy();
		expect(resolved.hiddenRibbonControls.has('view.show.selectionPane')).toBeTruthy();
		expect(resolved.hiddenRibbonControls.has('insert.links.link')).toBeTruthy();
		expect(resolved.hiddenElementCommands.has('hyperlink')).toBeTruthy();
		expect(resolved.keyboard.disabled.has('hyperlink')).toBeTruthy();
		expect(resolved.hiddenBackstageCards.has('protect')).toBeTruthy();
		expect(resolved.hiddenBackstagePages.has('options')).toBeTruthy();
		expect(resolved.hiddenRibbonControls.has('help.help.options')).toBeTruthy();
	});
});

describe('host File tab pages', () => {
	it('places host pages after the built-in pages of their group', () => {
		const mount = () => {};
		const resolved = resolveCustomization({
			backstage: {
				hiddenPages: ['close'],
				extraPages: [
					{ id: 'policies', label: 'Policies', render: mount },
					{ id: 'about', label: 'About', group: 'footer', render: 'x-about' },
					{ id: 'policies', label: 'Duplicate', render: mount },
				],
			},
		});
		const ids = backstageNavEntries(BACKSTAGE_NAV, resolved).map((entry) => entry.id);
		expect(ids).toStrictEqual([
			'home',
			'new',
			'open',
			'info',
			'save',
			'saveAs',
			'print',
			'share',
			'export',
			'host.policies',
			'account',
			'options',
			'host.about',
		]);
	});

	it('adds, replaces and removes host pages through the controller', () => {
		const { api, getResolved } = createCustomizationController();
		api.addBackstagePage({ id: 'a', label: 'A', render: 'x-a' });
		api.addBackstagePage({ id: 'a', label: 'A2', render: 'x-a' });
		expect(getResolved().hostBackstagePages.map((p) => p.label)).toStrictEqual(['A2']);
		api.removeBackstagePage('a');
		expect(getResolved().hostBackstagePages).toHaveLength(0);
		api.setRegion('statusBar', 'x-status');
		api.setCapability('edit', false);
		expect(isRegionReplaced(getResolved(), 'statusBar')).toBeTruthy();
		expect(isEditingAllowed(getResolved())).toBeFalsy();
		api.setRegion('statusBar', undefined);
		api.setCapability('edit', undefined);
		expect(isRegionReplaced(getResolved(), 'statusBar')).toBeFalsy();
		expect(isEditingAllowed(getResolved())).toBeTruthy();
	});
});
