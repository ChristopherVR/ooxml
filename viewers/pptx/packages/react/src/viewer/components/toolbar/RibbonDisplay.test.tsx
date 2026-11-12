// @vitest-environment happy-dom
/**
 * Ribbon Display Options in the desktop ribbon: the menu, "Show tabs only" (commands laid over
 * the slide by a chosen tab), full-screen mode (the reveal strip) and the host's `ribbonDisplay`
 * prop. The viewer's document gestures come from `useRibbonDisplayController`, so the harness
 * renders the ribbon the way the viewer root does.
 *
 * No `@testing-library/react` is available in this workspace, so this follows the manual
 * `createRoot` + `act` harness used elsewhere in the package.
 */
import type { RibbonDisplayMode } from 'ooxml-ui/pptx';
import React, { act, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { useRibbonDisplayController } from '../../hooks/useRibbonDisplayController';
import { Toolbar } from '../Toolbar';
import type { ToolbarProps } from './toolbar-types';

// The labels come back as their keys, so the menu's text is checkable.
vi.mock('../../i18n/use-translation', () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

const NOOP = (): void => undefined;

/** The ribbon's props for an edit-mode deck with nothing selected. Functions default to no-ops. */
function toolbarProps(): ToolbarProps {
	const base: Record<string, unknown> = {
		mode: 'edit',
		canEdit: true,
		isNarrowViewport: false,
		isSidebarCollapsed: false,
		isInspectorPaneOpen: false,
		isCompactToolbarOpen: false,
		toolbarSection: 'home',
		scale: 1,
		canUndo: false,
		canRedo: false,
		findReplaceOpen: false,
		selectedElement: null,
		selectedCount: 0,
		selectionGroupable: false,
		canDistribute: false,
		editTemplateMode: false,
		newShapeType: 'rect',
		activeTool: 'select',
		drawingColor: '#000000',
		drawingWidth: 2,
		clipboardPayload: null,
		layoutOptions: [],
		customShows: [],
		// Read by the tab click: it must be an own key, since the props are spread into the ribbon.
		onSetToolbarSection: NOOP,
		embeddedFontFamilies: [],
		themeFonts: [],
		customFontFamilies: [],
		hiddenActions: [],
		spellCheckEnabled: false,
		showGrid: false,
		showRulers: false,
		showGuides: false,
		snapToGrid: false,
		snapToShape: false,
		hasMacros: false,
		recentPresentationsCount: 0,
		slideCommentCount: 0,
		isCommentsPanelOpen: false,
	};
	// Any other callback the ribbon reads is a no-op, and only the ones it calls matter here.
	return new Proxy(base, {
		get: (target, key: string | symbol) => (key in target ? target[key as string] : NOOP),
	}) as unknown as ToolbarProps;
}

interface HarnessProps {
	ribbonDisplay?: RibbonDisplayMode;
	onRibbonDisplayChange?: (mode: RibbonDisplayMode) => void;
}

/** The viewer's shape: the controller comes from the hook and the root owns the gestures. */
function RibbonHarness({ ribbonDisplay, onRibbonDisplayChange }: HarnessProps): React.ReactElement {
	const rootRef = useRef<HTMLDivElement>(null);
	const ribbon = useRibbonDisplayController({
		ribbonDisplay,
		onRibbonDisplayChange,
		collapseRibbonAutomatically: false,
		rootRef,
	});
	return (
		<div ref={rootRef}>
			<Toolbar {...toolbarProps()} ribbonDisplay={ribbon} />
		</div>
	);
}

let container: HTMLDivElement;
let root: Root;
let actEnvironment: boolean | undefined;

beforeEach(() => {
	actEnvironment = globalThis.IS_REACT_ACT_ENVIRONMENT;
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = actEnvironment;
});

function render(props: HarnessProps): void {
	act(() => {
		root.render(<RibbonHarness {...props} />);
	});
}

function ribbonRoot(): HTMLElement {
	const element = container.querySelector<HTMLElement>('[data-pptx-chrome="ribbon"]');
	if (!element) {
		throw new Error('the ribbon did not render');
	}
	return element;
}

function ribbonContent(): HTMLElement {
	const element = container.querySelector<HTMLElement>('[data-pptx-chrome="ribbon-content"]');
	if (!element) {
		throw new Error('the ribbon content did not render');
	}
	return element;
}

function click(element: Element): void {
	act(() => {
		(element as HTMLElement).click();
	});
}

function tabByLabel(label: string): HTMLElement {
	const tab = Array.from(container.querySelectorAll<HTMLElement>('[role="tab"]')).find(
		(element) => element.textContent === label,
	);
	if (!tab) {
		throw new Error(`no tab ${label}`);
	}
	return tab;
}

function openMenu(): void {
	click(container.querySelector('[data-pptx-chrome="ribbon-display-options"]') as Element);
}

function menuItem(mode: RibbonDisplayMode): HTMLElement {
	const item = container.querySelector<HTMLElement>(
		`[role="menuitemradio"][data-ribbon-display="${mode}"]`,
	);
	if (!item) {
		throw new Error(`no menu item ${mode}`);
	}
	return item;
}

describe('ribbon display options in the Toolbar', () => {
	it('lists the three modes in PowerPoint order with the current one checked', () => {
		render({ ribbonDisplay: 'always' });
		openMenu();

		const items = Array.from(container.querySelectorAll<HTMLElement>('[role="menuitemradio"]'));
		expect(items.map((item) => item.textContent)).toEqual([
			'pptx.ribbon.display.fullScreen',
			'pptx.ribbon.display.tabs',
			'pptx.ribbon.display.always',
		]);
		expect(items.map((item) => item.getAttribute('aria-checked'))).toEqual([
			'false',
			'false',
			'true',
		]);
	});

	it('choosing Show tabs only hides the command content and reports the mode', () => {
		const onRibbonDisplayChange = vi.fn();
		render({ ribbonDisplay: 'always', onRibbonDisplayChange });
		expect(ribbonContent().classList.contains('hidden')).toBe(false);

		openMenu();
		click(menuItem('tabs'));

		expect(ribbonRoot().getAttribute('data-ribbon-display')).toBe('tabs');
		expect(ribbonContent().classList.contains('hidden')).toBe(true);
		expect(onRibbonDisplayChange).toHaveBeenCalledTimes(1);
		expect(onRibbonDisplayChange).toHaveBeenCalledWith('tabs');
	});

	it('a chosen tab lays its commands over the slide until Escape', () => {
		const onRibbonDisplayChange = vi.fn();
		render({ ribbonDisplay: 'tabs', onRibbonDisplayChange });
		expect(ribbonContent().classList.contains('hidden')).toBe(true);

		click(tabByLabel('pptx.ribbon.tab.home'));

		expect(ribbonRoot().hasAttribute('data-ribbon-peek')).toBe(true);
		expect(ribbonContent().classList.contains('hidden')).toBe(false);
		expect(ribbonContent().classList.contains('absolute')).toBe(true);
		expect(ribbonContent().classList.contains('bg-background')).toBe(true);

		act(() => {
			document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		});

		expect(ribbonRoot().hasAttribute('data-ribbon-peek')).toBe(false);
		expect(ribbonContent().classList.contains('hidden')).toBe(true);
		expect(onRibbonDisplayChange).not.toHaveBeenCalled();
	});

	it('a double click on a tab switches between the docked ribbon and the tab row', () => {
		const onRibbonDisplayChange = vi.fn();
		render({ ribbonDisplay: 'always', onRibbonDisplayChange });

		act(() => {
			tabByLabel('pptx.ribbon.tab.home').dispatchEvent(
				new MouseEvent('dblclick', { bubbles: true }),
			);
		});

		expect(onRibbonDisplayChange).toHaveBeenCalledWith('tabs');
	});

	it('starts in full-screen mode with the reveal strip and no ribbon tabs', () => {
		render({ ribbonDisplay: 'fullScreen' });

		const strip = container.querySelector<HTMLElement>('[data-pptx-chrome="ribbon-reveal"]');
		expect(strip).not.toBeNull();
		expect(ribbonRoot().getAttribute('data-ribbon-display')).toBe('fullScreen');
		expect(ribbonRoot().classList.contains('hidden')).toBe(true);
		expect(container.querySelector('[role="tab"]')).toBeNull();
	});

	it('pressing the strip in full-screen mode lays the ribbon over the slide, opaque', () => {
		render({ ribbonDisplay: 'fullScreen' });
		click(container.querySelector('[data-pptx-chrome="ribbon-reveal"]') as Element);

		expect(ribbonRoot().hasAttribute('data-ribbon-peek')).toBe(true);
		expect(ribbonRoot().classList.contains('absolute')).toBe(true);
		expect(ribbonRoot().classList.contains('bg-background')).toBe(true);
		// The strip keeps its place while the ribbon is peeked.
		expect(container.querySelector('[data-pptx-chrome="ribbon-reveal"]')).not.toBeNull();
	});

	it('a host change of the ribbonDisplay prop does not fire the change notification', () => {
		const onRibbonDisplayChange = vi.fn();
		render({ ribbonDisplay: 'always', onRibbonDisplayChange });

		render({ ribbonDisplay: 'tabs', onRibbonDisplayChange });

		expect(ribbonRoot().getAttribute('data-ribbon-display')).toBe('tabs');
		expect(ribbonContent().classList.contains('hidden')).toBe(true);
		expect(onRibbonDisplayChange).not.toHaveBeenCalled();
	});
});
