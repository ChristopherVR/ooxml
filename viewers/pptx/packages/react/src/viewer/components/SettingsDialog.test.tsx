// @vitest-environment happy-dom
import type { OfficeOptionCategory, ViewerOptions } from 'ooxml-ui/pptx';
import {
	cloneViewerOptions,
	DEFAULT_VIEWER_OPTIONS,
	resolveCustomization,
	VIEWER_OPTIONS_TABS,
} from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock(import('../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

const { SettingsDialog } = await import('./SettingsDialog');
const { ViewerCustomizationContext } = await import('./viewer-customization-context');
type SettingsDialogProps = import('./SettingsDialog').SettingsDialogProps;

type OptionsDialogElement = HTMLElement & {
	categories: readonly OfficeOptionCategory[];
	category: string;
	updateComplete: Promise<unknown>;
};

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

function createProps(overrides: Partial<SettingsDialogProps> = {}): SettingsDialogProps {
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

function renderDialog(props: SettingsDialogProps): void {
	act(() => {
		root.render(<SettingsDialog {...props} />);
	});
}

function dialog(): OptionsDialogElement {
	const element = container.querySelector<OptionsDialogElement>('pptx-ui-options-dialog');
	expect(element).not.toBeNull();
	return element!;
}

function shadowButton(action: string): HTMLButtonElement {
	return dialog().shadowRoot!.querySelector<HTMLButtonElement>(`[data-action="${action}"]`)!;
}

describe('settingsDialog', () => {
	it('renders nothing while closed', () => {
		renderDialog(createProps({ isOpen: false }));
		expect(container.innerHTML).toBe('');
	});

	it('hands all ten categories from VIEWER_OPTIONS_TABS to the shared dialog', () => {
		renderDialog(createProps());
		expect(VIEWER_OPTIONS_TABS).toHaveLength(10);
		expect(dialog().categories.map((entry) => entry.id)).toEqual(
			VIEWER_OPTIONS_TABS.map((tab) => tab.id),
		);
		expect(dialog().categories[0]?.label).toBe('pptx.settings.general');
	});

	it('leaves out the Language category when a language pick cannot apply', () => {
		renderDialog(createProps({ languageSelectable: false }));
		const ids = dialog().categories.map((entry) => entry.id);
		expect(ids).not.toContain('language');
		expect(ids).toHaveLength(VIEWER_OPTIONS_TABS.length - 1);
		renderDialog(createProps({ languageSelectable: true }));
		expect(dialog().categories.map((entry) => entry.id)).toContain('language');
	});

	it('hides the Language category on top of the pages the host hid', () => {
		const customization = resolveCustomization({ options: { hiddenPages: ['save'] } });
		act(() => {
			root.render(
				<ViewerCustomizationContext.Provider value={customization}>
					<SettingsDialog {...createProps({ languageSelectable: false })} />
				</ViewerCustomizationContext.Provider>,
			);
		});
		const ids = dialog().categories.map((entry) => entry.id);
		expect(ids).not.toContain('language');
		expect(ids).not.toContain('save');
		expect(ids).toHaveLength(VIEWER_OPTIONS_TABS.length - 2);
	});

	it('slots the product panes of the open category only', () => {
		renderDialog(createProps());
		expect(container.querySelector('[slot="block-themePicker"]')).not.toBeNull();
		expect(container.querySelector('[slot="pane-ribbon"]')).toBeNull();
		act(() => {
			dialog().category = 'ribbon';
		});
		expect(container.querySelector('[slot="pane-ribbon"]')).not.toBeNull();
		expect(container.querySelector('[slot="block-themePicker"]')).toBeNull();
	});

	it('reports a toggle change through onOptionChange', () => {
		const props = createProps();
		renderDialog(props);
		const checkbox = dialog().shadowRoot!.querySelector<HTMLInputElement>(
			'[data-key="general.showMiniToolbar"] input',
		);
		expect(checkbox).not.toBeNull();
		act(() => {
			checkbox!.click();
		});
		expect(props.onOptionChange).toHaveBeenCalledWith('general', 'showMiniToolbar', false);
	});

	it('shows the AI section only when aiEnabled is set', () => {
		renderDialog(createProps());
		expect(dialog().categories.some((entry) => entry.id === 'ai')).toBe(false);
		renderDialog(createProps({ aiEnabled: true }));
		expect(dialog().categories.at(-1)?.id).toBe('ai');
		act(() => {
			dialog().category = 'ai';
		});
		expect(container.querySelector('[slot="pane-ai"]')).not.toBeNull();
	});

	it('cancel restores the snapshot taken when the dialog opened', () => {
		const props = createProps();
		renderDialog(props);
		act(() => {
			shadowButton('cancel').click();
		});
		expect(props.onRestoreOptions).toHaveBeenCalledOnce();
		expect(vi.mocked(props.onRestoreOptions).mock.calls[0][0]).toStrictEqual(props.options);
		expect(props.onClose).toHaveBeenCalledOnce();
	});

	it('OK keeps the edits and closes; Reset goes through onResetOptions', () => {
		const props = createProps();
		renderDialog(props);
		act(() => {
			shadowButton('reset').click();
		});
		expect(props.onResetOptions).toHaveBeenCalledOnce();
		act(() => {
			shadowButton('ok').click();
		});
		expect(props.onClose).toHaveBeenCalledOnce();
		expect(props.onRestoreOptions).not.toHaveBeenCalled();
	});
});
