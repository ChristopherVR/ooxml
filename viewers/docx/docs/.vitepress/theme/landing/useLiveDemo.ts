import type { ComputedRef, Ref } from 'vue';
import { computed, onBeforeUnmount, onMounted, ref } from 'vue';
import { useData, withBase } from 'vitepress';

import { resolveViewerLocale } from './viewer-locale';

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

/** Languages the Word demo accepts through its `?locale=` parameter. */
export const DEMO_LOCALES: { id: string; label: string }[] = [
	{ id: 'en', label: 'English' },
	{ id: 'fr', label: 'Français' },
	{ id: 'de', label: 'Deutsch' },
	{ id: 'es', label: 'Español' },
	{ id: 'zh-CN', label: '简体中文' },
];

export type LiveDemoMode = 'solo' | 'collab';

export interface LiveDemoState {
	started: Ref<boolean>;
	mode: Ref<LiveDemoMode>;
	activeKey: Ref<string>;
	guestKey: Ref<string>;
	locale: Ref<string>;
	src: ComputedRef<string>;
	guestSrc: ComputedRef<string>;
	guestLabel: ComputedRef<string>;
	activeLabel: ComputedRef<string>;
	start: () => void;
	selectFramework: (key: string) => void;
	selectGuest: (key: string) => void;
	setLocale: (id: string) => void;
	setMode: (mode: LiveDemoMode) => void;
}

function frameworkByKey(key: string): DemoFramework {
	return DEMO_FRAMEWORKS.find((f) => f.key === key) ?? DEMO_FRAMEWORKS[0];
}

function randomRoom(): string {
	return `landing-${Math.random().toString(36).slice(2, 10)}`;
}

/**
 * State for the landing page's embedded live demo: a framework switcher over
 * the demo apps deployed beside the docs. Collaboration mode shows two windows,
 * each in a framework of its own, that join one session by name (a
 * BroadcastChannel of this browser, no network).
 *
 * The iframe only loads once the section scrolls near the viewport (or the
 * visitor clicks the load button), so visitors who never reach the section
 * download nothing.
 */
export function useLiveDemo(section: Ref<HTMLElement | null>): LiveDemoState {
	const started = ref(false);
	const mode = ref<LiveDemoMode>('solo');
	const activeKey = ref('react');
	const guestKey = ref('vue');
	// Starts in the site language (nav language menu); the visitor can still change it.
	const { lang } = useData();
	const locale = ref<string>(resolveViewerLocale(lang.value));

	/** The session name both windows join. A new one for every sharing session. */
	const room = ref(randomRoom());
	// Window A opens the sample and hosts the session; window B joins it by name and receives the
	// document. Each window picks its own framework: the session is a BroadcastChannel named after
	// the room, nothing more.
	const src = computed(() => {
		const { route } = frameworkByKey(activeKey.value);
		return withBase(
			mode.value === 'collab'
				? `/${route}/?sample=1&room=${room.value}&name=Ada&locale=${locale.value}`
				: `/${route}/?sample=1&locale=${locale.value}`,
		);
	});
	const guestSrc = computed(() =>
		withBase(
			`/${frameworkByKey(guestKey.value).route}/?room=${room.value}&name=Grace&locale=${locale.value}`,
		),
	);
	const guestLabel = computed(() => frameworkByKey(guestKey.value).label);
	const activeLabel = computed(() => frameworkByKey(activeKey.value).label);

	function start(): void {
		started.value = true;
	}

	function selectFramework(key: string): void {
		activeKey.value = key;
		started.value = true;
	}

	function selectGuest(key: string): void {
		guestKey.value = key;
	}

	function setLocale(id: string): void {
		locale.value = id;
		started.value = true;
	}

	function setMode(next: LiveDemoMode): void {
		mode.value = next;
		room.value = randomRoom();
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
		mode,
		activeKey,
		guestKey,
		locale,
		src,
		guestSrc,
		guestLabel,
		activeLabel,
		start,
		selectFramework,
		selectGuest,
		setLocale,
		setMode,
	};
}
