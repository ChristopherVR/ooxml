<script setup lang="ts">
import { withBase } from 'vitepress';

import { useLandingCopy } from './copy';

const copy = useLandingCopy();
</script>

<template>
	<section class="ot-section ot-faq" data-reveal>
		<div class="ot-bar">
			<h2 class="ot-h2">{{ copy.faq.title }}</h2>
			<p class="ot-bar__note">{{ copy.faq.note }}</p>
		</div>
		<div class="ot-faq__list">
			<details v-for="item in copy.faq.items" :key="item.q" class="ot-faq__item">
				<summary>{{ item.q }}</summary>
				<p>{{ item.a }}</p>
				<a v-if="item.href" :href="withBase(item.href)">{{ item.link }}</a>
			</details>
		</div>
	</section>
</template>

<style scoped>
.ot-faq__list {
	margin-top: 2rem;
	max-width: 52rem;
	border-top: 1px solid var(--ot-line);
}

.ot-faq__item {
	border-bottom: 1px solid var(--ot-line);
	padding: 1.1rem 0;
}

.ot-faq__item summary {
	cursor: pointer;
	font-family: var(--ot-display);
	font-size: 1.08rem;
	font-weight: 600;
}

.ot-faq__item p {
	margin: 0.8rem 0 0;
	line-height: 1.75;
	color: var(--ot-ink-soft);
}

.ot-faq__item a {
	display: inline-block;
	margin-top: 0.6rem;
	font-size: 0.8125rem;
	font-weight: 600;
	color: var(--ot-accent);
}
</style>

<style>
/* Open and close over 320 ms (docs/pages-design.md, Motion). Browsers that can animate the
   details content box do both directions; the rest fade the answer in on open. */
@media (prefers-reduced-motion: no-preference) {
	.ot-faq__item {
		interpolate-size: allow-keywords;
	}

	.ot-faq__item summary {
		transition: color 0.2s cubic-bezier(0.22, 1, 0.36, 1);
	}

	.ot-faq__item summary:hover {
		color: var(--ot-accent);
	}

	@supports selector(::details-content) {
		.ot-faq__item::details-content {
			block-size: 0;
			overflow: clip;
			opacity: 0;
			transition:
				block-size 320ms cubic-bezier(0.22, 1, 0.36, 1),
				opacity 320ms cubic-bezier(0.22, 1, 0.36, 1),
				content-visibility 320ms allow-discrete;
		}

		.ot-faq__item[open]::details-content {
			block-size: auto;
			opacity: 1;
		}
	}

	@supports not selector(::details-content) {
		.ot-faq__item[open] > :not(summary) {
			animation: ot-faq-open 320ms cubic-bezier(0.22, 1, 0.36, 1) backwards;
		}
	}
}

@keyframes ot-faq-open {
	from {
		opacity: 0;
		transform: translateY(-6px);
	}
}
</style>
