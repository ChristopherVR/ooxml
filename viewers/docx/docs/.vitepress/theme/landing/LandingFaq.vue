<script setup lang="ts">
import { withBase } from 'vitepress';

import { useLandingCopy } from './copy';

const copy = useLandingCopy();
</script>

<template>
	<section class="pv-section pv-faq">
		<header class="pv-sectionbar">
			<h2 class="pv-h2">{{ copy.faq.title }}</h2>
		</header>
		<div class="pv-faq__list">
			<details v-for="item in copy.faq.items" :key="item.q" class="pv-faq__item">
				<summary>
					<span>{{ item.q }}</span>
					<i aria-hidden="true">+</i>
				</summary>
				<p>
					{{ item.a }}
					<a v-if="item.link" :href="withBase(item.link.href)">{{ item.link.text }}</a>
				</p>
			</details>
		</div>
	</section>
</template>

<style scoped>
.pv-faq__list {
	margin-top: 2.2rem;
	max-width: 50rem;
	border-top: 1px solid var(--pv-line);
}

.pv-faq__item {
	border-bottom: 1px solid var(--pv-line);
}

.pv-faq__item summary {
	display: flex;
	align-items: center;
	justify-content: space-between;
	gap: 1.5rem;
	padding: 1.15rem 0.2rem;
	cursor: pointer;
	list-style: none;
	font-family: var(--pv-display);
	font-size: 1.02rem;
	font-weight: 600;
	letter-spacing: -0.01em;
}

.pv-faq__item summary::-webkit-details-marker {
	display: none;
}

.pv-faq__item summary:hover {
	color: var(--pv-accent);
}

.pv-faq__item summary i {
	font-style: normal;
	font-family: var(--pv-mono);
	color: var(--pv-accent);
	transition: transform 0.18s ease;
}

.pv-faq__item[open] summary i {
	transform: rotate(45deg);
}

.pv-faq__item p {
	margin: 0;
	padding: 0 0.2rem 1.3rem;
	max-width: 44rem;
	font-size: 0.92rem;
	line-height: 1.75;
	color: var(--pv-ink-soft);
}

.pv-faq__item p a {
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	color: var(--pv-accent);
	white-space: nowrap;
}
</style>
