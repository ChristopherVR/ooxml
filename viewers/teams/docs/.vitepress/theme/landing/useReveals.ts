import { onBeforeUnmount, onMounted, type Ref } from 'vue';

/**
 * Reveal-on-scroll for the landing page (docs/pages-design.md, Motion). Every `[data-reveal]`
 * element under `root` rises and fades in once, the first time 15% of it is visible. The hidden
 * state is only applied after mount (class `is-reveal` on the root), so the page stays readable
 * without JavaScript and under `prefers-reduced-motion: reduce`.
 */
export function useReveals(root: Ref<HTMLElement | null>): void {
	let observer: IntersectionObserver | null = null;

	onMounted(() => {
		const el = root.value;
		if (!el || typeof IntersectionObserver === 'undefined') return;
		if (window.matchMedia?.('(prefers-reduced-motion: reduce)').matches) return;
		observer = new IntersectionObserver(
			(entries) => {
				for (const entry of entries) {
					if (!entry.isIntersecting) continue;
					entry.target.classList.add('is-in');
					observer?.unobserve(entry.target);
				}
			},
			{ threshold: 0.15 },
		);
		el.classList.add('is-reveal');
		el.querySelectorAll('[data-reveal]').forEach((node) => observer?.observe(node));
	});

	onBeforeUnmount(() => {
		observer?.disconnect();
		observer = null;
		root.value?.classList.remove('is-reveal');
	});
}
