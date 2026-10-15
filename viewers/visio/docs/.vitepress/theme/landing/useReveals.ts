import type { Ref } from 'vue';
import { onBeforeUnmount, onMounted } from 'vue';

/**
 * Scroll reveals for the landing page. Elements marked `data-reveal` start
 * hidden (landing.css, only when motion is allowed) and rise in once as they
 * enter the viewport. Nothing re-animates when scrolling back up.
 */
export function useReveals(root: Ref<HTMLElement | null>): void {
	let observer: IntersectionObserver | null = null;

	onMounted(() => {
		const host = root.value;
		if (!host || typeof IntersectionObserver === 'undefined') {
			return;
		}
		observer = new IntersectionObserver(
			(entries) => {
				for (const entry of entries) {
					if (entry.isIntersecting) {
						entry.target.classList.add('is-revealed');
						observer?.unobserve(entry.target);
					}
				}
			},
			{ threshold: 0.15 },
		);
		for (const target of host.querySelectorAll('[data-reveal]')) {
			observer.observe(target);
		}
		// Hide only once the observer is watching, so no-JS visitors see everything.
		host.classList.add('pv-reveals');
	});

	onBeforeUnmount(() => {
		observer?.disconnect();
		observer = null;
	});
}
