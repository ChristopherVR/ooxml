<script setup lang="ts">
import { withBase } from 'vitepress';

import { useLandingCopy } from './copy';
import { featureIcon } from './icons';

const copy = useLandingCopy();
</script>

<template>
	<section class="pv-section pv-features">
		<header class="pv-sectionbar" data-reveal>
			<h2 class="pv-h2">{{ copy.features.title }}</h2>
		</header>
		<div class="pv-features__grid" data-reveal="grid">
			<article
				v-for="(item, index) in copy.features.items"
				:key="item.title"
				class="pv-features__card"
				:style="{ '--i': Math.min(index, 7) }"
			>
				<span class="pv-features__icon" aria-hidden="true">
					<svg
						viewBox="0 0 24 24"
						fill="none"
						stroke="currentColor"
						stroke-width="1.6"
						stroke-linecap="round"
						stroke-linejoin="round"
						v-html="featureIcon(item.icon, index)"
					></svg>
				</span>
				<span class="pv-features__go" aria-hidden="true">
					<svg
						viewBox="0 0 12 12"
						fill="none"
						stroke="currentColor"
						stroke-width="1.4"
						stroke-linecap="round"
						stroke-linejoin="round"
					>
						<path d="M3 9l6-6M4 3h5v5" />
					</svg>
				</span>
				<h3>{{ item.title }}</h3>
				<p>{{ item.copy }}</p>
				<a class="pv-features__link" :href="withBase(item.link.href)">{{ item.link.text }}</a>
			</article>
		</div>
	</section>
</template>

<style scoped>
.pv-features__grid {
	display: grid;
	grid-template-columns: repeat(4, minmax(0, 1fr));
	gap: 1px;
	background: var(--pv-line);
	border: 1px solid var(--pv-line);
	border-radius: 8px;
	overflow: hidden;
}

.pv-features__card {
	position: relative;
	display: flex;
	flex-direction: column;
	gap: 0.7rem;
	background: var(--pv-surface);
	padding: 1.2rem 1.2rem 1.4rem;
	transition:
		background-color 200ms var(--pv-ease),
		transform 200ms var(--pv-ease);
}

.pv-features__card:hover {
	background: var(--pv-accent-soft);
	z-index: 1;
}

@media (prefers-reduced-motion: no-preference) {
	.pv-features__card:hover {
		transform: translateY(-2px);
	}
}

.pv-features__icon {
	display: grid;
	place-items: center;
	width: 36px;
	height: 36px;
	margin-bottom: 0.4rem;
	border-radius: 6px;
	background: var(--pv-accent-tile);
	border: 1px solid var(--pv-accent-edge);
	color: var(--pv-accent);
}

.pv-features__icon svg {
	width: 20px;
	height: 20px;
}

.pv-features__go {
	position: absolute;
	top: 1.2rem;
	right: 1.2rem;
	display: grid;
	place-items: center;
	width: 22px;
	height: 22px;
	border: 1px solid var(--pv-line);
	border-radius: 4px;
	color: var(--pv-ink-soft);
	transition:
		color 200ms var(--pv-ease),
		border-color 200ms var(--pv-ease);
}

.pv-features__go svg {
	width: 10px;
	height: 10px;
}

.pv-features__card:hover .pv-features__go {
	color: var(--pv-accent);
	border-color: var(--pv-accent);
}

.pv-features__card h3 {
	font-family: var(--pv-display);
	font-size: 1rem;
	font-weight: 640;
	letter-spacing: -0.01em;
	line-height: 1.25;
	margin: 0;
}

.pv-features__card p {
	flex: 1;
	font-size: 0.86rem;
	line-height: 1.65;
	color: var(--pv-ink-soft);
	margin: 0;
}

.pv-features__link {
	font-family: var(--pv-ui);
	font-size: 0.8rem;
	font-weight: 600;
	color: var(--pv-accent);
	text-decoration: underline transparent;
	text-underline-offset: 3px;
	transition: text-decoration-color 200ms var(--pv-ease);
}

.pv-features__link:hover {
	text-decoration-color: currentColor;
}

/* The whole card is the link target */
.pv-features__link::after {
	content: '';
	position: absolute;
	inset: 0;
}

@media (max-width: 1100px) {
	.pv-features__grid {
		grid-template-columns: repeat(2, minmax(0, 1fr));
	}
}

@media (max-width: 560px) {
	.pv-features__grid {
		grid-template-columns: minmax(0, 1fr);
	}
}
</style>
