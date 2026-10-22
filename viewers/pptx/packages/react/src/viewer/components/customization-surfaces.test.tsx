// @vitest-environment happy-dom
/**
 * Host UI customisation at the React render sites that map a shared
 * descriptor: the File > Options dialog (hidden pages / settings, locked
 * settings) and the element context menu (hidden commands, disabled menu).
 * The decisions are `pptx-viewer-shared`'s; these
 * tests pin that React actually renders what they return.
 */
import type { ResolvedCustomization, ViewerCustomization, ViewerOptions } from 'ooxml-ui/pptx';
import { cloneViewerOptions, DEFAULT_VIEWER_OPTIONS, resolveCustomization } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock(import('../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

const { SettingsDialog } = await import('./SettingsDialog');
const { ContextMenu } = await import('./ContextMenu');
const { ViewerCustomizationContext } = await import('./viewer-customization-context');
type SettingsDialogProps = import('./SettingsDialog').SettingsDialogProps;
type ContextMenuProps = import('./context-menu-types').ContextMenuProps;

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

function renderWith(customization: ViewerCustomization, node: React.ReactElement): void {
	const resolved: ResolvedCustomization = resolveCustomization(customization);
	act(() => {
		root.render(
			<ViewerCustomizationContext.Provider value={resolved}>
				{node}
			</ViewerCustomizationContext.Provider>,
		);
	});
}

function dialogProps(overrides: Partial<SettingsDialogProps> = {}): SettingsDialogProps {
	return {
		isOpen: true,
		onClose: vi.fn<() => void>(),
		options: cloneViewerOptions(DEFAULT_VIEWER_OPTIONS),
		onOptionChange: vi.fn<SettingsDialogProps['onOptionChange']>(),
		onRestoreOptions: vi.fn<(options: ViewerOptions) => void>(),
		onRibbonTabHiddenChange: vi.fn<SettingsDialogProps['onRibbonTabHiddenChange']>(),
		onQuickAccessCommandsChange: vi.fn<(commandIds: string[]) => void>(),
		onResetOptions: vi.fn<SettingsDialogProps['onResetOptions']>(),
		onClearCache: vi.fn<() => void>(),
		themeKey: 'default',
		availableThemes: [],
		onSelectTheme: vi.fn<(key: string) => void>(),
		localeCode: 'en',
		availableLocales: [{ code: 'en', label: 'English', nativeLabel: 'English' }],
		onSelectLocale: vi.fn<(code: string) => void>(),
		...overrides,
	};
}

type OptionsDialogElement = HTMLElement & {
	categories: readonly {
		id: string;
		label: string;
		sections: readonly { controls: readonly { key: string; disabled?: string }[] }[];
	}[];
	category: string;
};

function optionsDialog(): OptionsDialogElement {
	return container.querySelector<OptionsDialogElement>('pptx-ui-options-dialog')!;
}

function navLabels(): string[] {
	return optionsDialog().categories.map((entry) => entry.label);
}

function allControls() {
	return optionsDialog().categories.flatMap((entry) =>
		entry.sections.flatMap((section) => section.controls),
	);
}

describe('file > Options customisation', () => {
	it('drops a hidden page from the rail and falls back to the first visible one', () => {
		renderWith(
			{ options: { hiddenPages: ['general', 'save'] } },
			<SettingsDialog {...dialogProps()} />,
		);
		expect(navLabels()).not.toContain('pptx.settings.general');
		expect(navLabels()).toContain('pptx.options.advanced.label');
		expect(optionsDialog().category).toBe(optionsDialog().categories[0]?.id);
	});

	it('hides the AI page when the host hid it, even with an assistant configured', () => {
		renderWith({}, <SettingsDialog {...dialogProps({ aiEnabled: true })} />);
		expect(navLabels()).toContain('pptx.ai.settingsSectionTitle');
		renderWith(
			{ options: { hiddenPages: ['ai'] } },
			<SettingsDialog {...dialogProps({ aiEnabled: true })} />,
		);
		expect(navLabels()).not.toContain('pptx.ai.settingsSectionTitle');
	});

	it('removes a hidden setting and renders a locked one disabled and inert', () => {
		const props = dialogProps();
		renderWith(
			{
				options: {
					hiddenSettings: ['general.userInitials'],
					locked: { 'general.userName': 'Host User' },
				},
			},
			<SettingsDialog {...props} />,
		);
		const controls = allControls();
		expect(controls.some((control) => control.key === 'general.userInitials')).toBe(false);
		const userName = controls.find((control) => control.key === 'general.userName');
		expect(userName?.disabled).toBe('pptx.options.lockedByHost');
		const field = optionsDialog().shadowRoot!.querySelector<HTMLInputElement>(
			'[data-key="general.userName"] input',
		);
		expect(field?.disabled).toBe(true);
		// An unlocked sibling stays enabled.
		expect(
			controls.find((control) => control.key === 'general.displayOptimization')?.disabled,
		).toBeUndefined();
		expect(props.onOptionChange).not.toHaveBeenCalled();
	});
});

function menuProps(): ContextMenuProps {
	return {
		contextMenuState: { x: 40, y: 60, elementId: 'sp_1' },
		mode: 'edit',
		selectedElement: null,
		tableEditorState: null,
		onAction: vi.fn(),
		onInsertTableRow: vi.fn(),
		onDeleteTableRow: vi.fn(),
		onInsertTableColumn: vi.fn(),
		onDeleteTableColumn: vi.fn(),
		onClose: vi.fn(),
	};
}

/** The rows of the shared context menu, drawn inside its open shadow root. */
function menuRows(): HTMLButtonElement[] {
	const host = container.querySelector('pptx-ui-context-menu');
	return Array.from(host?.shadowRoot?.querySelectorAll('button') ?? []);
}

describe('element context menu customisation', () => {
	it('runs a host element command with the effective selection and closes the menu', () => {
		const onSelect = vi.fn();
		const onClose = vi.fn();
		renderWith(
			{ contextMenu: { extraElementCommands: [{ id: 'chat', label: 'Send to chat', onSelect }] } },
			<ContextMenu {...menuProps()} slideIndex={2} elementIds={['a', 'b']} onClose={onClose} />,
		);
		const button = menuRows().find((item) => item.textContent === 'Send to chat')!;
		act(() => button.click());
		expect(onSelect).toHaveBeenCalledExactlyOnceWith({ slideIndex: 2, elementIds: ['a', 'b'] });
		expect(onClose).toHaveBeenCalledOnce();
	});

	function menuLabels(): string[] {
		return menuRows().map((button) => button.textContent ?? '');
	}

	it('omits a hidden command and keeps the rest', () => {
		renderWith({}, <ContextMenu {...menuProps()} />);
		expect(menuLabels()).toContain('pptx.contextMenu.delete');
		const before = menuLabels().length;
		renderWith(
			{ contextMenu: { hiddenElementCommands: ['delete'] } },
			<ContextMenu {...menuProps()} />,
		);
		expect(menuLabels()).not.toContain('pptx.contextMenu.delete');
		expect(menuLabels()).toHaveLength(before - 1);
	});

	it('renders nothing when the host disabled the menu', () => {
		renderWith({ contextMenu: { disableElementMenu: true } }, <ContextMenu {...menuProps()} />);
		expect(container.querySelector('[data-pptx-context-menu="true"]')).toBeNull();
		expect(container.innerHTML).toBe('');
	});
});
