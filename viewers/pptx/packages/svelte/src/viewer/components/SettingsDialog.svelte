<script lang="ts">
	import { onMount } from 'svelte';
	import {
		DEFAULT_QUICK_ACCESS_COMMAND_IDS,
		bindPptxOptionsDialog,
		resolveViewerAddinStatus,
	} from 'ooxml-ui/pptx';
	import type {
		PptxOptionsDialogBinding,
		PptxOptionsDialogElement,
		ThemeCatalogEntry,
		ViewerOptions,
	} from 'ooxml-ui/pptx';
	import type { LocaleCatalogEntry } from 'ooxml-ui/pptx/i18n';
	import { useTranslator } from '../../i18n/context';
	import { useViewerCustomization } from '../state/viewer-customization.svelte';
	import type { ViewerOptionsState } from '../state/viewer-options.svelte';
	import SettingsAiSection from './ai/SettingsAiSection.svelte';
	import SettingsAppearanceTab from './SettingsAppearanceTab.svelte';
	import SettingsCustomFontsSection from './SettingsCustomFontsSection.svelte';
	import SettingsLanguageTab from './SettingsLanguageTab.svelte';
	import OptionsAddInsPane from './settings/OptionsAddInsPane.svelte';
	import OptionsQuickAccessPane from './settings/OptionsQuickAccessPane.svelte';
	import OptionsRibbonPane from './settings/OptionsRibbonPane.svelte';

	/**
	 * File > Options: the shared `office-ui-options-dialog`, driven by `bindPptxOptionsDialog`
	 * (categories, settings, live apply, Cancel and Reset live in `ooxml-ui/pptx`). This only slots
	 * the panes PowerPoint has on top of the generic rows (themes, fonts, languages, Customize
	 * Ribbon, Quick Access Toolbar, Add-ins, AI) while their category is open.
	 */
	const {
		optionsState,
		onclose,
		themeKey,
		themeCatalog,
		onsetthemekey,
		locale,
		availableLocales,
		onsetlocale,
		aiEnabled = false,
		collabActive = false,
		customFontFamilies = [],
		oncustomfont = () => {},
	}: {
		/** The shared File > Options state (store + behavior projections). */
		optionsState: ViewerOptionsState;
		onclose: () => void;
		themeKey: string;
		themeCatalog: readonly ThemeCatalogEntry[];
		onsetthemekey: (key: string) => void;
		locale: string;
		availableLocales?: readonly LocaleCatalogEntry[];
		onsetlocale: (code: string) => void;
		/** When true, an "AI" section is shown for exporting detailed chat logs. */
		aiEnabled?: boolean;
		/** Live collaboration session state, for the Add-ins pane's status column. */
		collabActive?: boolean;
		/** Families registered this session via the Fonts section. */
		customFontFamilies?: readonly string[];
		/** A font file was registered; the ribbon adds the family to its list. */
		oncustomfont?: (family: string) => void;
	} = $props();

	const t = useTranslator();
	const customization = useViewerCustomization();
	const options = $derived<ViewerOptions>(optionsState.options);
	// Real runtime signals for the Add-ins pane's active/inactive split: the two
	// three.js-backed renderers follow Advanced > "Disable 3D rendering", and
	// the collaboration module follows the live session.
	const addinStatus = $derived(
		resolveViewerAddinStatus(options.advanced.disable3DRendering, collabActive),
	);

	let dialog = $state<PptxOptionsDialogElement | null>(null);
	let category = $state('general');
	let binding: PptxOptionsDialogBinding | null = null;

	onMount(() => {
		binding = bindPptxOptionsDialog(dialog!, {
			setOption: (group, key, value) => optionsState.setValue(group, key, value),
			restoreOptions: (snapshot) => optionsState.restore(snapshot),
			resetOptions: () => optionsState.reset(),
			close: onclose,
			categoryChanged: (id) => (category = id),
		});
		return () => {
			binding?.dispose();
			binding = null;
		};
	});
	const translate = (key: string): string => t(key);
	$effect(() => {
		// Re-labels on a language change and follows every option, customisation and AI flip.
		binding?.sync({
			options,
			t: translate,
			languageKey: locale,
			customization: customization.resolved,
			aiEnabled,
		});
	});

	function clearCache(): void {
		void optionsState.clearCache().catch(() => {
			// Keep the options dialog usable when IndexedDB rejects cleanup.
		});
	}
</script>

<pptx-ui-options-dialog bind:this={dialog}>
	{#if category === 'general'}
		<div slot="block-themePicker">
			<SettingsAppearanceTab {themeKey} {themeCatalog} onselect={onsetthemekey} />
		</div>
		<div slot="block-customFonts">
			<SettingsCustomFontsSection
				enabled={options.general.enableCustomFontUpload}
				families={customFontFamilies}
				onregistered={oncustomfont}
			/>
		</div>
	{:else if category === 'save'}
		<div slot="block-clearCache">
			<p class="hint">{t('pptx.options.save.clearCacheDescription')}</p>
			<button type="button" class="ghost" onclick={clearCache}>{t('pptx.options.save.clearCacheNow')}</button>
		</div>
	{:else if category === 'language'}
		<div slot="block-language">
			<SettingsLanguageTab {locale} {availableLocales} onselect={onsetlocale} />
		</div>
	{:else if category === 'ribbon'}
		<div slot="pane-ribbon">
			<OptionsRibbonPane {options} ontabhiddenchange={(id, hidden) => optionsState.setRibbonTabHidden(id, hidden)} onreset={() => optionsState.reset('ribbon')} />
		</div>
	{:else if category === 'quickAccess'}
		<div slot="pane-quickAccess">
			<OptionsQuickAccessPane {options} oncommandschange={(ids) => optionsState.setQuickAccessCommands(ids)} onreset={() => optionsState.setQuickAccessCommands([...DEFAULT_QUICK_ACCESS_COMMAND_IDS])} />
		</div>
	{:else if category === 'addIns'}
		<div slot="pane-addIns">
			<OptionsAddInsPane {addinStatus} />
		</div>
	{:else if category === 'ai'}
		<div slot="pane-ai">
			<SettingsAiSection />
		</div>
	{/if}
</pptx-ui-options-dialog>

<style>
	.hint { margin: 0 0 8px; color: var(--pptx-muted-foreground, #94a3b8); font-size: 11px; }
	button.ghost { border: 1px solid var(--pptx-border, #3f3f52); border-radius: 6px; padding: 6px 14px; background: transparent; color: var(--pptx-foreground, #e2e8f0); font-size: 12px; cursor: pointer; }
</style>
