<script setup lang="ts">
import { withBase } from 'vitepress';

import { useLandingCopy } from './copy';

const copy = useLandingCopy();

function resolve(href: string, external?: boolean): string {
	return external ? href : withBase(href);
}
</script>

<template>
	<section class="pv-finale">
		<div class="pv-finale__inner">
			<span class="pv-finale__mark" aria-hidden="true"><i></i></span>
			<h2 class="pv-finale__title">{{ copy.finale.title }}</h2>
			<p class="pv-finale__sub">{{ copy.finale.sub }}</p>
			<div class="pv-finale__actions">
				<a class="pv-btn pv-btn--accent" :href="withBase(copy.finale.quick.href)">
					<span>{{ copy.finale.quick.text }}</span>
				</a>
				<a
					class="pv-btn pv-btn--outline"
					href="https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio"
				>
					<span>{{ copy.finale.github }}</span>
				</a>
			</div>
		</div>
		<div class="pv-finale__columns">
			<div v-for="column in copy.finale.columns" :key="column.title" class="pv-finale__column">
				<span class="pv-finale__coltitle">{{ column.title }}</span>
				<a
					v-for="link in column.links"
					:key="link.text"
					:href="resolve(link.href, link.external)"
					:target="link.external ? '_blank' : undefined"
					:rel="link.external ? 'noreferrer' : undefined"
				>
					{{ link.text }}
				</a>
			</div>
		</div>
		<div class="pv-finale__foot">
			<span>{{ copy.finale.bottomLeft }}</span>
			<span>{{ copy.finale.bottomRight }}</span>
		</div>
	</section>
</template>

<style scoped>
.pv-finale {
	background: #181a1e;
	color: #f0f0f0;
	--pv-accent: #7b93dc;
	--pv-on-accent: #181a1e;
}

.dark .pv-finale {
	background: #202228;
	border-top: 1px solid var(--pv-line);
}

.pv-finale__inner {
	max-width: 84rem;
	margin: 0 auto;
	padding: clamp(6rem, 16vh, 11rem) clamp(1.4rem, 5vw, 4.5rem) clamp(4rem, 9vh, 6.5rem);
	text-align: center;
}

.pv-finale__mark {
	display: inline-grid;
	place-items: center;
	width: 3.2rem;
	height: 3.2rem;
	background: var(--pv-accent);
	border-radius: 6px;
	margin-bottom: 2rem;
}

.pv-finale__mark i {
	width: 1.5rem;
	height: calc(1.5rem * 9 / 16);
	border: 2px solid #fff;
	border-radius: 2px;
}

.pv-finale__title {
	font-family: var(--pv-display);
	font-size: clamp(2.3rem, 5.6vw, 4.6rem);
	font-weight: 660;
	line-height: 1.02;
	letter-spacing: -0.02em;
	margin: 1.4rem auto 1.5rem;
	max-width: 20ch;
}

.pv-finale__sub {
	color: rgba(240, 240, 240, 0.62);
	line-height: 1.75;
	max-width: 36rem;
	margin: 0 auto;
}

.pv-finale__actions {
	display: flex;
	justify-content: center;
	flex-wrap: wrap;
	gap: 0.9rem;
	margin-top: 2.8rem;
}

.pv-btn--accent {
	background: var(--pv-accent);
	border: 1px solid var(--pv-accent);
	color: var(--pv-on-accent);
}

.pv-btn--accent:hover {
	background: #8ba0e0;
	border-color: #8ba0e0;
}

.pv-btn--outline {
	border: 1px solid rgba(240, 240, 240, 0.35);
	color: #f0f0f0;
}

.pv-btn--outline:hover {
	background: rgba(240, 240, 240, 0.08);
	border-color: #f0f0f0;
}

.pv-finale__columns {
	max-width: 84rem;
	margin: 0 auto;
	padding: 3rem clamp(1.4rem, 5vw, 4.5rem);
	border-top: 1px solid rgba(240, 240, 240, 0.12);
	display: grid;
	grid-template-columns: repeat(3, minmax(0, 1fr));
	gap: 2rem;
}

.pv-finale__column {
	display: flex;
	flex-direction: column;
	gap: 0.65rem;
}

.pv-finale__coltitle {
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	color: rgba(240, 240, 240, 0.45);
	margin-bottom: 0.4rem;
}

.pv-finale__column a {
	font-size: 0.88rem;
	color: rgba(240, 240, 240, 0.72);
	transition: color 0.2s ease;
	width: fit-content;
}

.pv-finale__column a:hover {
	color: var(--pv-accent);
}

.pv-finale__foot {
	max-width: 84rem;
	margin: 0 auto;
	padding: 1.4rem clamp(1.4rem, 5vw, 4.5rem) 2.4rem;
	border-top: 1px solid rgba(240, 240, 240, 0.12);
	display: flex;
	justify-content: space-between;
	flex-wrap: wrap;
	gap: 0.8rem;
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	color: rgba(240, 240, 240, 0.45);
}

@media (max-width: 700px) {
	.pv-finale__columns {
		grid-template-columns: 1fr;
	}
}
</style>
