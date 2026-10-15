import type {
	PptxOptionsDialogElement,
	ResolvedCustomization,
	ThemeCatalogEntry,
	ViewerAddinStatus,
	ViewerOptionsStore,
	ViewerOptionsTabId,
	ViewerTheme,
} from 'ooxml-ui/pptx';
import {
	AI_OPTIONS_CATEGORY,
	bindPptxOptionsDialog,
	EMPTY_RESOLVED_CUSTOMIZATION,
	optionsBlockSlot,
	optionsPaneSlot,
	registerPptxWebControls,
} from 'ooxml-ui/pptx';
import type { PptxAiChatStore } from 'ooxml-ui/pptx/ai';
import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';

import { createAiSettingsSection } from '../ai/ai-settings-section';
import type { Translator } from '../i18n';
import { createEl } from '../render';
import { appendOptionsAction } from './options/options-controls';
import { createAddinsPaneState, renderAddInsPane } from './options/options-pane-addins';
import {
	createQuickAccessPaneState,
	renderQuickAccessPane,
} from './options/options-pane-quick-access';
import { renderRibbonPane } from './options/options-pane-ribbon';
import { ownOverlay } from './overlay-owner';
import { appendRadioRow, appendSwatchRow } from './parity-dialog-shell';
import type { RegionSet } from './region-set';
import { appendCustomFontsPane } from './settings-custom-fonts';

/** File > Options > General appearance wiring: theme catalog + active key + selection sink. */
export interface SettingsDialogThemeOptions {
	catalog: readonly ThemeCatalogEntry[];
	currentKey: string;
	onSelect: (theme: ViewerTheme | undefined, key: string) => void;
}

/** File > Options > Language wiring: locale catalog + active code + selection sink. */
export interface SettingsDialogLocaleOptions {
	catalog: readonly LocaleCatalogEntry[];
	currentCode: string;
	onSelect: (code: string) => void;
}

export interface ViewerOptionsDialogDeps {
	/** The shared File > Options store; changes apply live, Cancel restores. */
	store: ViewerOptionsStore;
	themeOptions: SettingsDialogThemeOptions;
	localeOptions: SettingsDialogLocaleOptions;
	/** Availability flags for the Add-ins pane (unset ids default to active). */
	addinStatus?: ViewerAddinStatus;
	/** Options > Save > "Delete cached files". */
	onClearCache: () => void;
	initialTab?: ViewerOptionsTabId;
	/** When true, an "AI" section is appended for exporting detailed chat logs. */
	aiEnabled?: boolean;
	/** The host's resolved UI customisation: hidden pages/sections/settings, locks. */
	customization?: ResolvedCustomization;
	/** Chat store the AI export reads from (defaults to the shared store). */
	aiChatStore?: PptxAiChatStore;
	/** The viewer's dialog regions: a host `dialog.options` renderer fills the dialog body. */
	regions?: RegionSet;
	/**
	 * Session-scoped custom font registry backing the Fonts section.
	 *
	 * Optional so a host embedding the dialog on its own can leave the feature
	 * out entirely; the pane then shows the list as empty.
	 */
	customFonts?: {
		list(): readonly string[];
		register(family: string): void;
	};
}

/**
 * The PowerPoint "File > Options" dialog: the shared `office-ui-options-dialog` driven by
 * `bindPptxOptionsDialog` (vanilla counterpart of React's `SettingsDialog`). The categories,
 * settings, live apply, Cancel and Reset are the shared layer's; this only fills the slots of the
 * panes PowerPoint has on top of the generic rows (themes, fonts, languages, Customize Ribbon,
 * Quick Access Toolbar, Add-ins, AI) for the open category.
 *
 * `t` must be a LIVE translator (the caller passes a delegating wrapper, see
 * `parity-workflows`): picking a language on the Language pane re-labels the whole open dialog
 * (title, category rail, footer, pane) under the new locale, as the other bindings' dialogs do.
 */
export function openSettingsDialog(
	doc: Document,
	t: Translator,
	deps: ViewerOptionsDialogDeps,
): void {
	registerPptxWebControls();
	const { store } = deps;
	const dialog = doc.createElement('pptx-ui-options-dialog') as PptxOptionsDialogElement;
	const customization = deps.customization ?? EMPTY_RESOLVED_CUSTOMIZATION;
	let category: string = deps.initialTab ?? 'general';
	let selectedThemeKey = deps.themeOptions.currentKey;
	let selectedLocaleCode = deps.localeOptions.currentCode;
	const quickAccessState = createQuickAccessPaneState();
	const addinsState = createAddinsPaneState();

	const slot = (name: string): HTMLElement => {
		const host = createEl(doc, 'div');
		host.slot = name;
		dialog.appendChild(host);
		return host;
	};

	/** Rebuild the product panes of the open category inside their slots. */
	const renderSlots = (): void => {
		// The host's `dialog.options` body (slot "body") belongs to the shared binding.
		for (const child of [...dialog.children]) {
			if (child.slot !== 'body') {
				child.remove();
			}
		}
		switch (category) {
			case 'general': {
				appendSwatchRow(
					doc,
					slot(optionsBlockSlot('themePicker')),
					deps.themeOptions.catalog.map((entry) => ({
						key: entry.key,
						label: t(entry.labelKey),
						previewColor: entry.theme?.colors?.primary ?? '#6b7280',
					})),
					selectedThemeKey,
					(key) => {
						const entry = deps.themeOptions.catalog.find((candidate) => candidate.key === key);
						if (entry) {
							selectedThemeKey = key;
							deps.themeOptions.onSelect(entry.theme, key);
							renderSlots();
						}
					},
				);
				appendCustomFontsPane(doc, t, slot(optionsBlockSlot('customFonts')), {
					enabled: store.getOptions().general.enableCustomFontUpload,
					families: deps.customFonts?.list() ?? [],
					onRegistered: (family) => deps.customFonts?.register(family),
					onRefresh: renderSlots,
				});
				break;
			}
			case 'save': {
				const host = slot(optionsBlockSlot('clearCache'));
				const description = createEl(doc, 'p', 'pptxv-options-section-desc');
				description.textContent = t('pptx.options.save.clearCacheDescription');
				host.appendChild(description);
				appendOptionsAction(doc, host, t('pptx.options.save.clearCacheNow'), deps.onClearCache);
				break;
			}
			case 'language':
				appendRadioRow(
					doc,
					slot(optionsBlockSlot('language')),
					'pptxv-options-locale',
					deps.localeOptions.catalog.map((entry) => ({
						value: entry.code,
						label: entry.nativeLabel,
					})),
					selectedLocaleCode,
					(code) => {
						selectedLocaleCode = code;
						deps.localeOptions.onSelect(code);
						// The locale just changed under the open dialog: every label in it is stale.
						sync();
						renderSlots();
					},
				);
				break;
			case 'ribbon':
				renderRibbonPane(doc, t, slot(optionsPaneSlot('ribbon')), store);
				break;
			case 'quickAccess':
				renderQuickAccessPane(
					doc,
					t,
					slot(optionsPaneSlot('quickAccess')),
					store,
					quickAccessState,
					renderSlots,
				);
				break;
			case 'addIns':
				renderAddInsPane(
					doc,
					t,
					slot(optionsPaneSlot('addIns')),
					deps.addinStatus,
					addinsState,
					renderSlots,
				);
				break;
			case AI_OPTIONS_CATEGORY:
				slot(optionsPaneSlot(AI_OPTIONS_CATEGORY)).appendChild(
					createAiSettingsSection({ doc, t, store: deps.aiChatStore }),
				);
				break;
			default:
				break;
		}
	};

	const support = deps.regions?.support;
	// The dialog (and a host body in it) goes away with the viewer.
	const close = ownOverlay(deps.regions, (): void => {
		unsubscribe();
		offRegion?.();
		binding.dispose();
		dialog.remove();
	});
	const binding = bindPptxOptionsDialog(
		dialog,
		{
			setOption: (group, key, value) => store.setValue(group, key, value),
			restoreOptions: (snapshot) => store.setOptions(snapshot),
			resetOptions: () => store.reset(),
			close,
			categoryChanged: (id) => {
				category = id;
				renderSlots();
			},
		},
		category,
	);
	const sync = (): void =>
		binding.sync({
			options: store.getOptions(),
			t,
			languageKey: selectedLocaleCode,
			customization,
			aiEnabled: deps.aiEnabled === true,
			regionContext: support?.buildContext('dialog.options', close) ?? null,
		});
	// Live re-render on any store commit (a control edit, reset, or an external change).
	const unsubscribe = store.subscribe(() => {
		if (!dialog.isConnected) {
			unsubscribe();
			return;
		}
		sync();
		renderSlots();
	});
	// A host body follows the viewer too (selection, slide, mode), not only the options.
	const offRegion = support?.subscribe(() => {
		if (dialog.isConnected) sync();
	});
	doc.body.appendChild(dialog);
	sync();
}
