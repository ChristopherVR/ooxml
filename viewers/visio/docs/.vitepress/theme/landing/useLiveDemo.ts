import type { ComputedRef, Ref } from 'vue';
import { computed, onBeforeUnmount, onMounted, ref } from 'vue';
import { withBase } from 'vitepress';

/**
 * One embeddable framework demo, built by scripts/build-pages.mjs and placed
 * next to the docs on GitHub Pages: `/demo/` (vanilla, the long-standing route), `/demo-react/`, `/demo-vue/`, etc.
 */
export interface DemoFramework {
	key: string;
	label: string;
	route: string;
}

export const DEMO_FRAMEWORKS: DemoFramework[] = [
	{ key: 'react', label: 'React', route: 'demo-react' },
	{ key: 'vue', label: 'Vue', route: 'demo-vue' },
	{ key: 'angular', label: 'Angular', route: 'demo-angular' },
	{ key: 'svelte', label: 'Svelte', route: 'demo-svelte' },
	{ key: 'solid', label: 'Solid', route: 'demo-solid' },
	{ key: 'vanilla', label: 'Vanilla JS', route: 'demo' },
];

export type LiveDemoMode = 'solo' | 'collab';

export interface LiveDemoState {
	started: Ref<boolean>;
	mode: Ref<LiveDemoMode>;
	activeKey: Ref<string>;
	guestKey: Ref<string>;
	src: ComputedRef<string>;
	activeLabel: ComputedRef<string>;
	start: () => void;
	selectFramework: (key: string) => void;
	selectGuest: (key: string) => void;
	setMode: (mode: LiveDemoMode) => void;
}

function frameworkByKey(key: string): DemoFramework {
	return DEMO_FRAMEWORKS.find((f) => f.key === key) ?? DEMO_FRAMEWORKS[0];
}

/** Pick a guest framework different from the host so the pairing shows cross-adapter use. */
function fallbackGuest(hostKey: string): string {
	return DEMO_FRAMEWORKS.find((f) => f.key !== hostKey)?.key ?? hostKey;
}

/**
 * State for the landing page's embedded live demo: a framework switcher over
 * the demo apps deployed beside the docs. Sharing mode shows two windows of the
 * same demo; the visitor starts File > Share in both (a BroadcastChannel session,
 * no network), so nothing is started on the visitor's behalf.
 *
 * The iframe only loads once the section scrolls near the viewport (or the
 * visitor clicks the load button), so visitors who never reach the section
 * download nothing.
 */
export function useLiveDemo(section: Ref<HTMLElement | null>): LiveDemoState {
	const started = ref(false);
	const mode = ref<LiveDemoMode>('solo');
	const activeKey = ref('vanilla');
	const guestKey = ref('vue');

	const src = computed(() => withBase(`/${frameworkByKey(activeKey.value).route}/?sample=1`));
	const activeLabel = computed(() => frameworkByKey(activeKey.value).label);

	function start(): void {
		started.value = true;
	}

	function selectFramework(key: string): void {
		activeKey.value = key;
		if (guestKey.value === key) {
			guestKey.value = fallbackGuest(key);
		}
		started.value = true;
	}

	function selectGuest(key: string): void {
		guestKey.value = key === activeKey.value ? fallbackGuest(activeKey.value) : key;
	}

	function setMode(next: LiveDemoMode): void {
		mode.value = next;
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
		src,
		activeLabel,
		start,
		selectFramework,
		selectGuest,
		setMode,
	};
}
