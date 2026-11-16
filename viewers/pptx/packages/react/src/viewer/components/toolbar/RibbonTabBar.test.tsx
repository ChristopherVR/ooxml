// @vitest-environment happy-dom
import { DEFAULT_VIEWER_OPTIONS, ribbonTab, ribbonTabButtons } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { ViewerOptionsContext } from '../viewer-options-context';
import type { RibbonTabBarProps } from './RibbonTabBar';
import { RibbonTabBar } from './RibbonTabBar';

vi.mock(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => key.replace(/^pptx\.ribbon\.tab\./u, '') }),
}));

const containers: HTMLElement[] = [];
afterEach(() => {
	for (const container of containers.splice(0)) container.remove();
});

function mount(props: Partial<RibbonTabBarProps> = {}, hiddenTabIds: string[] = []) {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	const container = document.createElement('div');
	document.body.append(container);
	containers.push(container);
	const root = createRoot(container);
	const all: RibbonTabBarProps = {
		activeSection: 'home',
		contextualTabs: [],
		activeContextual: null,
		onSelectSection: vi.fn(),
		onSelectContextual: vi.fn(),
		...props,
	};
	const options = {
		...DEFAULT_VIEWER_OPTIONS,
		ribbon: { ...DEFAULT_VIEWER_OPTIONS.ribbon, hiddenTabIds },
	} as typeof DEFAULT_VIEWER_OPTIONS;
	const render = (next: Partial<RibbonTabBarProps> = {}) =>
		act(() =>
			root.render(
				<ViewerOptionsContext.Provider value={options}>
					<RibbonTabBar {...all} {...next}>
						<button type='button'>Share</button>
					</RibbonTabBar>
				</ViewerOptionsContext.Provider>,
			),
		);
	render();
	return { container, render, props: all, unmount: () => act(() => root.unmount()) };
}

const labels = (root: ParentNode) => ribbonTabButtons(root).map((tab) => tab.textContent);

describe('RibbonTabBar', () => {
	it('draws every fixed tab on the shared tab row, the actions in its slot', () => {
		const { container, unmount } = mount();
		const host = container.querySelector('pptx-ui-ribbon-tabs')!;
		expect(host.getAttribute('data-pptx-chrome')).toBe('ribbon-tabs');
		expect(labels(container)).toEqual([
			'file',
			'home',
			'insert',
			'draw',
			'design',
			'transitions',
			'animations',
			'slideShow',
			'record',
			'review',
			'view',
			'help',
		]);
		expect(ribbonTab(container, '[aria-selected="true"]')!.textContent).toBe('home');
		expect(host.querySelector('[slot="actions"]')!.textContent).toBe('Share');
		unmount();
	});

	it('leaves out tabs hidden by the host or by Customize Ribbon', () => {
		const { container, unmount } = mount({ hiddenActions: ['review'] }, ['draw', 'file']);
		expect(labels(container)).not.toContain('review');
		expect(labels(container)).not.toContain('draw');
		expect(labels(container)).toContain('file');
		unmount();
	});

	it('routes a chosen tab by its kind and a double click to the collapse toggle', () => {
		const onToggleCollapsed = vi.fn();
		const { container, props, render, unmount } = mount({
			contextualTabs: ['shapeFormat'],
			addInTabs: [{ id: 'tools', label: 'Tools', groups: [] }],
			onSelectAddIn: vi.fn(),
			onToggleCollapsed,
		});
		act(() => ribbonTab(container, '[data-tab="insert"]')!.click());
		act(() => ribbonTab(container, '[data-ribbon-contextual-tab="shapeFormat"]')!.click());
		act(() => ribbonTab(container, '[data-ribbon-add-in-tab="tools"]')!.click());
		expect(props.onSelectSection).toHaveBeenCalledWith('insert');
		expect(props.onSelectContextual).toHaveBeenCalledWith('shapeFormat');
		expect(props.onSelectAddIn).toHaveBeenCalledWith('tools');
		render({ onSelectAddIn: props.onSelectAddIn, activeContextual: 'shapeFormat' });
		expect(ribbonTab(container, '[aria-selected="true"]')!.dataset.tab).toBe('shapeFormat');
		ribbonTab(container, '[data-tab="view"]')!.dispatchEvent(
			new MouseEvent('dblclick', { bubbles: true, composed: true }),
		);
		expect(onToggleCollapsed).toHaveBeenCalledOnce();
		unmount();
	});
});
