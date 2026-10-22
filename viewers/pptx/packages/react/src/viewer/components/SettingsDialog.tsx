import type {
	PptxOptionsDialogBinding,
	PptxOptionsDialogElement,
	ThemeCatalogEntry,
	ToolbarTabId,
	ViewerAddinStatus,
	ViewerOptions,
	ViewerOptionsGroupId,
} from 'ooxml-ui/pptx';
import {
	AI_OPTIONS_CATEGORY,
	DEFAULT_QUICK_ACCESS_COMMAND_IDS,
	bindPptxOptionsDialog,
	optionsBlockSlot,
	optionsPaneSlot,
} from 'ooxml-ui/pptx';
import type { PptxAiChatStore } from 'ooxml-ui/pptx/ai';
import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';
import React, { useContext, useLayoutEffect, useMemo, useRef, useState } from 'react';

import { useTranslation } from '../i18n/use-translation';
import { OptionsAddInsPane } from './settings/OptionsAddInsPane';
import { OptionsQuickAccessPane } from './settings/OptionsQuickAccessPane';
import { OptionsRibbonPane } from './settings/OptionsRibbonPane';
import { SettingsAiTab } from './SettingsAiTab';
import { SettingsAppearanceTab } from './SettingsAppearanceTab';
import { SettingsCustomFontsSection } from './SettingsCustomFontsSection';
import { SettingsLanguageTab } from './SettingsLanguageTab';
import { useViewerCustomizationContext } from './viewer-customization-context';
import { ViewerRegionBaseContext } from './ViewerRegion';

/**
 * Stable empty default for `customFontFamilies`. A fresh `[]` in the parameter
 * list is a new reference every render, which breaks referential equality for
 * anything downstream that memoises on it.
 */
const NO_CUSTOM_FONT_FAMILIES: readonly string[] = [];

export interface SettingsDialogProps {
	isOpen: boolean;
	onClose: () => void;
	/** Full File > Options snapshot rendered by every pane. */
	options: ViewerOptions;
	onOptionChange: (
		group: ViewerOptionsGroupId,
		key: string,
		value: boolean | number | string,
	) => void;
	/** Restore a snapshot wholesale (Cancel semantics). */
	onRestoreOptions: (options: ViewerOptions) => void;
	onRibbonTabHiddenChange: (tabId: ToolbarTabId, hidden: boolean) => void;
	onQuickAccessCommandsChange: (commandIds: string[]) => void;
	onResetOptions: (group?: ViewerOptionsGroupId) => void;
	/** Options > Save > "Delete cached files". */
	onClearCache: () => void;
	/** Availability flags for the Add-ins pane. */
	addinStatus?: ViewerAddinStatus;
	/** Key of the currently active theme catalog entry. */
	themeKey: string;
	availableThemes: readonly ThemeCatalogEntry[];
	onSelectTheme: (key: string) => void;
	/** Currently active locale code. */
	localeCode: string;
	availableLocales: readonly LocaleCatalogEntry[];
	onSelectLocale: (code: string) => void;
	/** When true, an "AI" section is shown for exporting detailed chat logs. */
	aiEnabled?: boolean;
	/** Families registered this session via the Fonts section. */
	customFontFamilies?: readonly string[];
	/** Called after a font file is registered, so the ribbon picks it up. */
	onCustomFontRegistered?: (family: string) => void;
	/** Chat store the AI section reads from (defaults to the shared store). */
	chatStore?: PptxAiChatStore;
}

/**
 * File > Options: the shared `office-ui-options-dialog` driven by
 * `bindPptxOptionsDialog` (categories, settings, live apply, Cancel and Reset all live in
 * `ooxml-ui/pptx`). This binding only supplies the panes PowerPoint has on top of the generic
 * rows (themes, fonts, languages, Customize Ribbon, Quick Access Toolbar, Add-ins, AI), slotted
 * into the element while their category is open.
 */
export function SettingsDialog(props: SettingsDialogProps): React.ReactElement | null {
	return props.isOpen ? <OpenSettingsDialog {...props} /> : null;
}

function OpenSettingsDialog({
	onClose,
	options,
	onOptionChange,
	onRestoreOptions,
	onRibbonTabHiddenChange,
	onQuickAccessCommandsChange,
	onResetOptions,
	onClearCache,
	addinStatus,
	themeKey,
	availableThemes,
	onSelectTheme,
	localeCode,
	availableLocales,
	onSelectLocale,
	aiEnabled,
	chatStore,
	customFontFamilies = NO_CUSTOM_FONT_FAMILIES,
	onCustomFontRegistered,
}: SettingsDialogProps): React.ReactElement {
	const { t } = useTranslation();
	const customization = useViewerCustomizationContext();
	const regionBase = useContext(ViewerRegionBaseContext);
	// A host `dialog.options` region fills the body; the shared binding adds the options model.
	const regionContext = useMemo(
		() =>
			regionBase ? { ...regionBase, region: 'dialog.options' as const, close: onClose } : null,
		[regionBase, onClose],
	);
	const dialogRef = useRef<PptxOptionsDialogElement>(null);
	const bindingRef = useRef<PptxOptionsDialogBinding | null>(null);
	const [category, setCategory] = useState('general');
	// The element outlives renders; its listeners always call the latest callbacks.
	const latest = useRef({ onClose, onOptionChange, onRestoreOptions, onResetOptions });
	latest.current = { onClose, onOptionChange, onRestoreOptions, onResetOptions };

	useLayoutEffect(() => {
		const binding = bindPptxOptionsDialog(dialogRef.current!, {
			setOption: (group, key, value) => latest.current.onOptionChange(group, key, value),
			restoreOptions: (snapshot) => latest.current.onRestoreOptions(snapshot),
			resetOptions: () => latest.current.onResetOptions(),
			close: () => latest.current.onClose(),
			categoryChanged: setCategory,
		});
		bindingRef.current = binding;
		return () => {
			binding.dispose();
			bindingRef.current = null;
		};
	}, []);
	useLayoutEffect(() => {
		bindingRef.current?.sync({
			options,
			t,
			customization,
			aiEnabled: Boolean(aiEnabled),
			regionContext,
		});
	}, [options, t, customization, aiEnabled, regionContext]);

	return (
		<pptx-ui-options-dialog ref={dialogRef}>
			{category === 'general' && (
				<>
					<div slot={optionsBlockSlot('themePicker')} className='mt-2'>
						<SettingsAppearanceTab
							activeThemeKey={themeKey}
							themes={availableThemes}
							onSelectTheme={onSelectTheme}
						/>
					</div>
					<div slot={optionsBlockSlot('customFonts')}>
						<SettingsCustomFontsSection
							enabled={options.general.enableCustomFontUpload}
							families={customFontFamilies}
							onRegistered={(family) => onCustomFontRegistered?.(family)}
						/>
					</div>
				</>
			)}
			{category === 'save' && (
				<div slot={optionsBlockSlot('clearCache')} className='mt-2'>
					<p className='mb-2 text-xs text-muted-foreground'>
						{t('pptx.options.save.clearCacheDescription')}
					</p>
					<button
						type='button'
						onClick={onClearCache}
						className='rounded border border-border px-3 py-1.5 text-xs text-foreground transition-colors hover:bg-accent'
					>
						{t('pptx.options.save.clearCacheNow')}
					</button>
				</div>
			)}
			{category === 'language' && (
				<div slot={optionsBlockSlot('language')}>
					<SettingsLanguageTab
						activeLocale={localeCode}
						locales={availableLocales}
						onSelectLocale={onSelectLocale}
					/>
				</div>
			)}
			{category === 'ribbon' && (
				<div slot={optionsPaneSlot('ribbon')}>
					<OptionsRibbonPane
						options={options}
						onRibbonTabHiddenChange={onRibbonTabHiddenChange}
						onResetRibbon={() => onResetOptions('ribbon')}
					/>
				</div>
			)}
			{category === 'quickAccess' && (
				<div slot={optionsPaneSlot('quickAccess')}>
					<OptionsQuickAccessPane
						options={options}
						onQuickAccessCommandsChange={onQuickAccessCommandsChange}
						onResetQuickAccess={() =>
							onQuickAccessCommandsChange([...DEFAULT_QUICK_ACCESS_COMMAND_IDS])
						}
					/>
				</div>
			)}
			{category === 'addIns' && (
				<div slot={optionsPaneSlot('addIns')}>
					<OptionsAddInsPane addinStatus={addinStatus} />
				</div>
			)}
			{category === AI_OPTIONS_CATEGORY && (
				<div slot={optionsPaneSlot(AI_OPTIONS_CATEGORY)}>
					<SettingsAiTab store={chatStore} />
				</div>
			)}
		</pptx-ui-options-dialog>
	);
}
