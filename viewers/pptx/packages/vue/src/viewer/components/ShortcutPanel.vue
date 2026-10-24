<script setup lang="ts">
/**
 * ShortcutPanel - grouped, searchable keyboard-shortcut help overlay.
 *
 * Lists the shared `VIEWER_SHORTCUT_REFERENCE`, the keymap every binding's help
 * shows, under its shared headings (`groupShortcutReference`), with a live
 * filter input and the platform's modifier (⌘ or Ctrl for `Ctrl/Cmd`). It
 * reuses {@link ModalDialog} for the teleported shell, backdrop/Esc close, and
 * header chrome.
 *
 * Presentational: the parent owns the `open` flag and handles `close`.
 */
import { groupShortcutReference, VIEWER_SHORTCUT_REFERENCE } from 'ooxml-ui/pptx';
import { computed, ref, watch } from 'vue';
import { useI18n } from 'vue-i18n';

import ModalDialog from './ModalDialog.vue';

const props = defineProps<{
	/** Whether the help overlay is visible. */
	open: boolean;
}>();

const emit = defineEmits<{
	(e: 'close'): void;
}>();

const { t } = useI18n();

/** Live filter query over the shortcut descriptions + combos. */
const query = ref('');

// Reset the filter each time the panel opens so it never reopens pre-filtered.
watch(
	() => props.open,
	(isOpen) => {
		if (isOpen) {
			query.value = '';
		}
	},
);

/** True on macOS: drives the ⌘ vs Ctrl glyph. */
const isMac = computed(() => {
	if (typeof navigator === 'undefined') {
		return false;
	}
	const platform = navigator.platform ?? '';
	const ua = navigator.userAgent ?? '';
	return /mac|iphone|ipad|ipod/iu.test(platform) || /mac/iu.test(ua);
});

/** A shortcut as this platform writes it: `Ctrl/Cmd` is ⌘ on macOS and Ctrl elsewhere. */
function platformShortcut(shortcut: string): string {
	return shortcut.replaceAll('Ctrl/Cmd', isMac.value ? '⌘' : 'Ctrl');
}

/** The reference filtered by the live query, then grouped for display. */
const groups = computed(() => {
	const needle = query.value.trim().toLowerCase();
	const matches =
		needle.length === 0
			? VIEWER_SHORTCUT_REFERENCE
			: VIEWER_SHORTCUT_REFERENCE.filter(
					(item) =>
						t(item.actionKey).toLowerCase().includes(needle) ||
						platformShortcut(item.shortcut).toLowerCase().includes(needle) ||
						item.shortcut.toLowerCase().includes(needle),
				);
	return groupShortcutReference(matches);
});

const hasResults = computed(() => groups.value.length > 0);

function requestClose(): void {
	emit('close');
}
</script>

<template>
	<ModalDialog :open="open" :title="t('pptx.settings.keyboardShortcuts')" @close="requestClose">
		<div
			class="pptx-vue-shortcuts flex min-w-[18rem] flex-col gap-3"
			data-pptx-shortcuts-panel="true"
		>
			<input
				v-model="query"
				type="text"
				class="pptx-vue-shortcuts-search w-full rounded border border-border bg-muted px-2.5 py-1.5 text-foreground outline-none focus:ring-1 focus:ring-primary"
				:placeholder="t('pptx.shortcuts.searchPlaceholder')"
				:aria-label="t('pptx.shortcuts.searchLabel')"
			/>

			<div
				v-if="hasResults"
				class="pptx-vue-shortcuts-list flex max-h-[22rem] flex-col gap-4 overflow-y-auto"
			>
				<section
					v-for="bucket in groups"
					:key="bucket.group"
					class="pptx-vue-shortcuts-group flex flex-col gap-1"
				>
					<h3
						class="pptx-vue-shortcuts-group-title text-[11px] font-semibold uppercase tracking-wide text-muted-foreground"
					>
						{{ t(bucket.labelKey) }}
					</h3>
					<div
						v-for="shortcut in bucket.shortcuts"
						:key="shortcut.actionKey"
						class="pptx-vue-shortcuts-row flex items-center justify-between gap-3 rounded bg-muted px-2 py-1.5"
					>
						<span class="pptx-vue-shortcuts-desc text-[13px] text-foreground">
							{{ t(shortcut.actionKey) }}
						</span>
						<kbd
							class="pptx-vue-shortcuts-combo shrink-0 whitespace-nowrap rounded border border-border bg-popover px-2 py-0.5 font-mono text-xs text-foreground"
						>
							{{ platformShortcut(shortcut.shortcut) }}
						</kbd>
					</div>
				</section>
			</div>

			<p v-else class="pptx-vue-shortcuts-empty px-2 py-3 text-[13px] text-muted-foreground">
				{{ t('pptx.shortcuts.noResults', { query }) }}
			</p>
		</div>
	</ModalDialog>
</template>
