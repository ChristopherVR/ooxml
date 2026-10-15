<script setup lang="ts">
import { onContentUpdated } from 'vitepress';

/**
 * Replays the page entrance (custom.css, `.pv-page-enter`) each time the docs content changes:
 * the article rises 12 px and fades in over 650 ms. Renderless; mounted once from Layout.vue.
 */
onContentUpdated(() => {
	if (typeof document === 'undefined') return;
	if (matchMedia('(prefers-reduced-motion: reduce)').matches) return;
	for (const el of document.querySelectorAll<HTMLElement>('.VPDoc .content, .VPDocAside')) {
		el.classList.remove('pv-page-enter');
		void el.offsetWidth;
		el.classList.add('pv-page-enter');
	}
});
</script>

<template><span hidden /></template>
