<script setup lang="ts">
import { useData, useRouter, withBase } from 'vitepress';
import { computed } from 'vue';

/**
 * The head of the sidebar on wide screens, where it stands in for the top navigation bar
 * (custom.css hides the bar there): the product, the appearance switch, the search field, the
 * site's navigation entries as shortcuts and the language. Same layout as the documentation hub.
 * Narrow screens keep VitePress's own header and never show this.
 */
const props = defineProps<{
	/** The product's letter on its colour tile, as in the suite (W, X, P, V, T). */
	letter: string;
}>();

const { site, theme, isDark, localeIndex, page } = useData();
const router = useRouter();

const EXTERNAL = /^https?:/u;

/** The navigation bar's own entries, so each locale's labels come along. */
const shortcuts = computed(() =>
	((theme.value.nav ?? []) as { text?: string; link?: string }[]).flatMap((item) =>
		item.text && item.link
			? [
					{
						text: item.text,
						link: EXTERNAL.test(item.link) ? item.link : withBase(item.link),
					},
				]
			: [],
	),
);

const ICONS = [
	// Guide: an open book.
	'M4 5h6a2 2 0 0 1 2 2v12a2 2 0 0 0-2-2H4zM20 5h-6a2 2 0 0 0-2 2v12a2 2 0 0 1 2-2h6z',
	// The suite: an app window.
	'M3 6a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2zM3 9h18',
	// Releases: a tag.
	'M3 4h8l10 10-7 7L3 10zM7.5 8.5h.01',
];

const locales = computed(() =>
	Object.entries(site.value.locales ?? {}).map(([key, locale]) => ({
		key,
		label: locale.label ?? key,
	})),
);

/** The page being read, in another language: the same path under that locale's folder. */
function localeHref(key: string): string {
	const current = localeIndex.value === 'root' ? '' : `${localeIndex.value}/`;
	const path = page.value.relativePath
		.slice(current.length)
		.replace(/(^|\/)index\.md$/u, '$1')
		.replace(/\.md$/u, '');
	return withBase(`/${key === 'root' ? '' : `${key}/`}${path}`);
}

function changeLocale(event: Event): void {
	void router.go(localeHref((event.target as HTMLSelectElement).value));
}

/** VitePress owns the search dialog (it stays mounted in the hidden bar); its shortcut opens it. */
function openSearch(): void {
	window.dispatchEvent(new KeyboardEvent('keydown', { key: 'k', ctrlKey: true, bubbles: true }));
}

const home = computed(() =>
	withBase(localeIndex.value === 'root' ? '/' : `/${localeIndex.value}/`),
);
const searchLabel = computed(
	() =>
		(theme.value.search?.options?.translations?.button?.buttonText as string | undefined) ??
		'Search',
);
/** The repository link the top bar shows as an icon. */
const source = computed(
	() =>
		((theme.value.socialLinks ?? []) as { icon?: unknown; link?: string }[]).find(
			(item) => item.icon === 'github' && item.link,
		)?.link,
);
const appearanceLabel = computed(
	() => (theme.value.darkModeSwitchLabel as string | undefined) ?? 'Appearance',
);
</script>

<template>
	<div class="pv-side-top">
		<div class="pv-side-brand-row">
			<a class="pv-side-brand" :href="home">
				<span class="pv-side-badge" aria-hidden="true">{{ props.letter }}</span>
				<span class="pv-side-name">{{ site.title }}</span>
			</a>
			<a v-if="source" class="pv-side-icon" :href="source" aria-label="GitHub" title="GitHub">
				<svg viewBox="0 0 24 24" aria-hidden="true">
					<path d="m8 7-5 5 5 5M16 7l5 5-5 5M14 4l-4 16" />
				</svg>
			</a>
			<button
				type="button"
				class="pv-side-icon"
				:aria-label="appearanceLabel"
				:title="appearanceLabel"
				:aria-pressed="isDark"
				@click="isDark = !isDark"
			>
				<svg viewBox="0 0 24 24" aria-hidden="true">
					<path d="M12 3a9 9 0 1 0 9 9 7 7 0 0 1-9-9z" />
				</svg>
			</button>
		</div>
		<button type="button" class="pv-side-search" @click="openSearch">
			<svg viewBox="0 0 24 24" aria-hidden="true">
				<circle cx="10" cy="10" r="6" />
				<path d="m15 15 6 6" />
			</svg>
			<span>{{ searchLabel }}</span>
			<kbd>/</kbd>
		</button>
		<div class="pv-side-shortcuts">
			<a v-for="(item, index) in shortcuts" :key="item.link" :href="item.link">
				<svg viewBox="0 0 24 24" aria-hidden="true"><path :d="ICONS[index % ICONS.length]" /></svg>
				{{ item.text }}
			</a>
		</div>
		<label v-if="locales.length > 1" class="pv-side-locale">
			<svg viewBox="0 0 24 24" aria-hidden="true">
				<circle cx="12" cy="12" r="9" />
				<path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18" />
			</svg>
			<select
				:value="localeIndex"
				:aria-label="theme.langMenuLabel ?? 'Language'"
				@change="changeLocale"
			>
				<option v-for="locale in locales" :key="locale.key" :value="locale.key">
					{{ locale.label }}
				</option>
			</select>
		</label>
	</div>
</template>
