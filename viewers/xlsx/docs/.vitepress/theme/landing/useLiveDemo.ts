import type { ComputedRef, Ref } from 'vue';
import { computed, onBeforeUnmount, onMounted, ref } from 'vue';
import { useData, withBase } from 'vitepress';

/**
 * One embeddable framework demo, built by scripts/build-pages.mjs and placed
 * next to the docs on GitHub Pages: `/demo/` (React), `/demo-vue/`, etc.
 */
export interface DemoFramework {
	key: string;
	label: string;
	route: string;
}

export const DEMO_FRAMEWORKS: DemoFramework[] = [
	{ key: 'react', label: 'React', route: 'demo' },
	{ key: 'vue', label: 'Vue', route: 'demo-vue' },
	{ key: 'angular', label: 'Angular', route: 'demo-angular' },
	{ key: 'svelte', label: 'Svelte', route: 'demo-svelte' },
	{ key: 'solid', label: 'Solid', route: 'demo-solid' },
	{ key: 'vanilla', label: 'Vanilla JS', route: 'demo-vanilla' },
];

const LOCALE_BY_LANG: Record<string, string> = {
	'en-US': 'en',
	'fr-FR': 'fr',
	'de-DE': 'de',
	'es-ES': 'es',
	'zh-CN': 'zh-CN',
};

/** Maps the VitePress site language (e.g. `'fr-FR'`) to a demo locale id. */
export function resolveDemoLocale(lang: string): string {
	return LOCALE_BY_LANG[lang] ?? 'en';
}

/** Languages the Excel demo accepts through its `?locale=` parameter. */
export const DEMO_LOCALES: { id: string; label: string }[] = [
	{ id: 'en', label: 'English' },
	{ id: 'fr', label: 'Français' },
	{ id: 'de', label: 'Deutsch' },
	{ id: 'es', label: 'Español' },
	{ id: 'zh-CN', label: '简体中文' },
];

export interface LiveDemoState {
	started: Ref<boolean>;
	activeKey: Ref<string>;
	locale: Ref<string>;
	src: ComputedRef<string>;
	fullSrc: ComputedRef<string>;
	activeLabel: ComputedRef<string>;
	start: () => void;
	selectFramework: (key: string) => void;
	setLocale: (id: string) => void;
}

function frameworkByKey(key: string): DemoFramework {
	return DEMO_FRAMEWORKS.find((f) => f.key === key) ?? DEMO_FRAMEWORKS[0];
}

/**
 * State for the landing page's embedded live demo: a framework switcher over
 * the demo apps deployed beside the docs. The embed opens the sample workbook
 * (`?sample=1`).
 *
 * The iframe only loads once the section scrolls near the viewport (or the
 * visitor clicks the load button), so visitors who never reach the section
 * download nothing.
 */
export function useLiveDemo(section: Ref<HTMLElement | null>): LiveDemoState {
	const started = ref(false);
	const activeKey = ref('react');
	// Defaults to the site language; the visitor can still change it.
	const { lang } = useData();
	const locale = ref(resolveDemoLocale(lang.value));

	const src = computed(() =>
		withBase(`/${frameworkByKey(activeKey.value).route}/?sample=1&locale=${locale.value}`),
	);
	const fullSrc = computed(() =>
		withBase(`/${frameworkByKey(activeKey.value).route}/?locale=${locale.value}`),
	);
	const activeLabel = computed(() => frameworkByKey(activeKey.value).label);

	function start(): void {
		started.value = true;
	}

	function selectFramework(key: string): void {
		activeKey.value = key;
		started.value = true;
	}

	function setLocale(id: string): void {
		locale.value = id;
		started.value = true;
	}

	let observer: IntersectionObserver | null = null;
	onMounted(() => {
		if (started.value || !section.value || typeof IntersectionObserver === 'undefined') {
			return;
		}
		observer = new IntersectionObserver(
			(entries) => {
				if (entries.some((entry) => entry.isIntersecting)) {
					started.value = true;
					observer?.disconnect();
					observer = null;
				}
			},
			{ rootMargin: '300px 0px' },
		);
		observer.observe(section.value);
	});
	onBeforeUnmount(() => {
		observer?.disconnect();
		observer = null;
	});

	return {
		started,
		activeKey,
		locale,
		src,
		fullSrc,
		activeLabel,
		start,
		selectFramework,
		setLocale,
	};
}
