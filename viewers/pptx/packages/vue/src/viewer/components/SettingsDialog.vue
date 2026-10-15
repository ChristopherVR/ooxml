<script setup lang="ts">
/**
 * SettingsDialog - the PowerPoint "File > Options" dialog.
 *
 * Vue counterpart of the React `SettingsDialog.tsx`: the shared
 * `office-ui-options-dialog`, driven by `bindPptxOptionsDialog` (the ten categories, their
 * settings, live apply, Cancel and Reset all live in `ooxml-ui/pptx`). This only slots the
 * panes PowerPoint has on top of the generic rows (themes, fonts, languages, Customize Ribbon,
 * Quick Access Toolbar, Add-ins, AI) while their category is open.
 */
import type {
	PptxOptionsDialogBinding,
	PptxOptionsDialogElement,
	ToolbarTabId,
	ViewerAddinStatus,
	ViewerOptions,
	ViewerOptionsGroupId,
} from 'ooxml-ui/pptx';
import {
	AI_OPTIONS_CATEGORY,
	bindPptxOptionsDialog,
	DEFAULT_QUICK_ACCESS_COMMAND_IDS,
	optionsBlockSlot,
	optionsPaneSlot,
} from 'ooxml-ui/pptx';
import type { PptxAiChatStore } from 'ooxml-ui/pptx/ai';
import { nextTick, onBeforeUnmount, ref, watch } from 'vue';
import { useI18n } from 'vue-i18n';

import type { LocaleCatalogEntry } from '../../i18n';
import type { ThemeCatalogEntry } from '../../theme';
import { useResolvedCustomization } from '../composables/useViewerCustomization';
import OptionsAddInsPane from './settings/OptionsAddInsPane.vue';
import OptionsQuickAccessPane from './settings/OptionsQuickAccessPane.vue';
import OptionsRibbonPane from './settings/OptionsRibbonPane.vue';
import SettingsAiTab from './SettingsAiTab.vue';
import SettingsAppearanceTab from './SettingsAppearanceTab.vue';
import SettingsCustomFontsSection from './SettingsCustomFontsSection.vue';
import SettingsLanguageTab from './SettingsLanguageTab.vue';

const props = defineProps<{
	/** Whether the dialog is visible. */
	open: boolean;
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
	/** Key of the currently active theme-catalog entry. */
	themeKey: string;
	/** Invoked with the entry's `key` when the user picks an Appearance swatch. */
	onThemeSelect: (key: string) => void;
	/** Code of the currently active locale. */
	localeCode: string;
	/** Invoked with the entry's `code` when the user picks a Language choice. */
	onLocaleSelect: (code: string) => void;
	/** Theme choices to offer; defaults to the built-in catalog. */
	availableThemes?: ThemeCatalogEntry[];
	/** Locale choices to offer on the Language pane. */
	availableLocales: LocaleCatalogEntry[];
	/** When true, an "AI" section is shown for exporting detailed chat logs. */
	aiEnabled?: boolean;
	/** Families registered this session via the Fonts section. */
	customFontFamilies?: readonly string[];
	/** Chat store the AI section reads from (defaults to the shared store). */
	chatStore?: PptxAiChatStore;
}>();

const emit = defineEmits<{
	/** Emitted when the dialog should close (changes kept). */
	(e: 'close'): void;
	/** A font file was registered; the ribbon adds the family to its list. */
	(e: 'customFontRegistered', family: string): void;
}>();

const { t, locale } = useI18n();
const resolved = useResolvedCustomization();
const dialog = ref<PptxOptionsDialogElement | null>(null);
const category = ref('general');
let binding: PptxOptionsDialogBinding | null = null;

function unbind(): void {
	binding?.dispose();
	binding = null;
}

const translate = (key: string): string => t(key);

function sync(): void {
	binding?.sync({
		options: props.options,
		t: translate,
		languageKey: locale.value,
		customization: resolved.value,
		aiEnabled: Boolean(props.aiEnabled),
	});
}

// The element exists only while the dialog is open; bind it when it mounts.
watch(
	dialog,
	(element) => {
		unbind();
		if (!element) {
			return;
		}
		binding = bindPptxOptionsDialog(element, {
			setOption: (group, key, value) => props.onOptionChange(group, key, value),
			restoreOptions: (snapshot) => props.onRestoreOptions(snapshot),
			resetOptions: () => props.onResetOptions(),
			close: () => emit('close'),
			categoryChanged: (id) => {
				category.value = id;
			},
		});
		sync();
	},
	{ flush: 'post' },
);
watch(
	() => [props.options, locale.value, resolved.value, props.aiEnabled],
	() => nextTick(sync),
);
onBeforeUnmount(unbind);

function resetQuickAccess(): void {
	props.onQuickAccessCommandsChange([...DEFAULT_QUICK_ACCESS_COMMAND_IDS]);
}
</script>

<template>
	<Teleport to="body">
		<pptx-ui-options-dialog v-if="open" ref="dialog">
			<template v-if="category === 'general'">
				<div :slot="optionsBlockSlot('themePicker')" class="mt-2">
					<SettingsAppearanceTab
						:theme-key="themeKey"
						:catalog="availableThemes"
						:on-select="onThemeSelect"
					/>
				</div>
				<div :slot="optionsBlockSlot('customFonts')">
					<SettingsCustomFontsSection
						:enabled="props.options.general.enableCustomFontUpload"
						:families="props.customFontFamilies ?? []"
						@registered="(family: string) => emit('customFontRegistered', family)"
					/>
				</div>
			</template>
			<div v-else-if="category === 'save'" :slot="optionsBlockSlot('clearCache')" class="mt-2">
				<p class="mb-2 text-xs text-muted-foreground">
					{{ t('pptx.options.save.clearCacheDescription') }}
				</p>
				<button
					type="button"
					class="rounded border border-border px-3 py-1.5 text-xs text-foreground transition-colors hover:bg-accent"
					@click="onClearCache"
				>
					{{ t('pptx.options.save.clearCacheNow') }}
				</button>
			</div>
			<div v-else-if="category === 'language'" :slot="optionsBlockSlot('language')">
				<SettingsLanguageTab
					:locale-code="localeCode"
					:locales="availableLocales"
					:on-select="onLocaleSelect"
				/>
			</div>
			<div v-else-if="category === 'ribbon'" :slot="optionsPaneSlot('ribbon')">
				<OptionsRibbonPane
					:options="options"
					:on-ribbon-tab-hidden-change="onRibbonTabHiddenChange"
					:on-reset-ribbon="() => onResetOptions('ribbon')"
				/>
			</div>
			<div v-else-if="category === 'quickAccess'" :slot="optionsPaneSlot('quickAccess')">
				<OptionsQuickAccessPane
					:options="options"
					:on-quick-access-commands-change="onQuickAccessCommandsChange"
					:on-reset-quick-access="resetQuickAccess"
				/>
			</div>
			<div v-else-if="category === 'addIns'" :slot="optionsPaneSlot('addIns')">
				<OptionsAddInsPane :addin-status="addinStatus" />
			</div>
			<div
				v-else-if="category === AI_OPTIONS_CATEGORY"
				:slot="optionsPaneSlot(AI_OPTIONS_CATEGORY)"
			>
				<SettingsAiTab :store="chatStore" />
			</div>
		</pptx-ui-options-dialog>
	</Teleport>
</template>
