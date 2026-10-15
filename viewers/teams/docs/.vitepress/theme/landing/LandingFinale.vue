<script setup lang="ts">
import { withBase } from 'vitepress';

import { useLandingCopy } from './copy';

const copy = useLandingCopy();
</script>

<template>
	<section class="ot-section ot-finale" data-reveal>
		<h2 class="ot-h2">{{ copy.finale.title }}</h2>
		<p class="ot-copy">
			{{ copy.finale.introBefore }}<code>npx openteams-server</code>{{ copy.finale.introAfter }}
		</p>
		<div class="ot-finale__actions">
			<a class="ot-btn ot-btn--solid" :href="withBase(copy.finale.start.href)">{{
				copy.finale.start.text
			}}</a>
			<a
				class="ot-btn ot-btn--ghost"
				href="https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams"
				target="_blank"
				rel="noopener"
				>{{ copy.finale.github }}</a
			>
		</div>
		<div class="ot-finale__cols">
			<nav v-for="col in copy.finale.columns" :key="col.title" :aria-label="col.title">
				<h3>{{ col.title }}</h3>
				<ul>
					<li v-for="link in col.links" :key="link.href">
						<a
							:href="link.external ? link.href : withBase(link.href)"
							:target="link.external ? '_blank' : link.href === '/demo/' ? '_self' : undefined"
							:rel="link.external ? 'noopener' : undefined"
							>{{ link.text }}</a
						>
					</li>
				</ul>
			</nav>
		</div>
		<p class="ot-finale__bottom">
			<span>{{ copy.finale.license }}</span>
			<span>{{ copy.finale.disclaimer }}</span>
		</p>
	</section>
</template>

<style scoped>
.ot-copy code {
	font-family: var(--ot-mono);
	font-size: 0.86em;
}

.ot-finale__actions {
	display: flex;
	flex-wrap: wrap;
	gap: 0.8rem;
	margin-top: 1.8rem;
}

.ot-finale__cols {
	display: grid;
	grid-template-columns: repeat(3, minmax(0, 1fr));
	gap: 2rem;
	margin-top: 4rem;
	padding-top: 2.4rem;
	border-top: 1px solid var(--ot-line);
}

.ot-finale__cols h3 {
	margin: 0 0 0.9rem;
	font-size: 0.8125rem;
	font-weight: 600;
	color: var(--ot-accent);
}

.ot-finale__cols ul {
	list-style: none;
	margin: 0;
	padding: 0;
	display: grid;
	gap: 0.55rem;
}

.ot-finale__cols a {
	color: var(--ot-ink-soft);
	font-size: 0.92rem;
}

.ot-finale__cols a:hover {
	color: var(--ot-accent);
}

.ot-finale__bottom {
	display: flex;
	flex-wrap: wrap;
	justify-content: space-between;
	gap: 0.6rem 2rem;
	margin: 3rem 0 0;
	font-size: 0.75rem;
	color: var(--ot-ink-soft);
}

@media (max-width: 700px) {
	.ot-finale__cols {
		grid-template-columns: minmax(0, 1fr);
	}
}
</style>
