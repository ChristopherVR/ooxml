<script setup lang="ts">
import { ref } from 'vue';
import { withBase } from 'vitepress';

import { useLandingCopy } from './copy';
import { icons } from './icons';

import LandingFaq from './LandingFaq.vue';
import LandingFinale from './LandingFinale.vue';
import LandingHero from './LandingHero.vue';
import LandingLiveDemo from './LandingLiveDemo.vue';
import LandingQuickstart from './LandingQuickstart.vue';
import { useReveals } from './useReveals';

const copy = useLandingCopy();
const root = ref<HTMLElement | null>(null);
useReveals(root);
</script>

<template>
	<div ref="root" class="ot-landing">
		<LandingHero />

		<LandingLiveDemo />

		<section class="ot-section" data-reveal>
			<div class="ot-bar">
				<h2 class="ot-h2">{{ copy.embed.title }}</h2>
				<p class="ot-bar__note">{{ copy.embed.note }}</p>
			</div>
			<div class="ot-grid ot-grid--two">
				<article
					v-for="(demo, i) in copy.embed.demos"
					:key="demo.href"
					class="ot-card"
					data-reveal
					:style="{ '--i': i }"
				>
					<h3>{{ demo.title }}</h3>
					<p>{{ demo.copy }}</p>
					<a class="ot-more" :href="withBase(demo.href)" target="_self">{{
						copy.embed.openDemo
					}}</a>
				</article>
			</div>
		</section>

		<section class="ot-section" data-reveal>
			<div class="ot-bar">
				<h2 class="ot-h2">{{ copy.features.title }}</h2>
				<p class="ot-bar__note">{{ copy.features.note }}</p>
			</div>
			<div class="ot-grid">
				<a
					v-for="(item, i) in copy.features.items"
					:key="item.title"
					class="ot-card ot-tool"
					data-reveal
					:style="{ '--i': i }"
					:href="withBase(item.href)"
					:aria-label="`${item.title}: ${item.link}`"
				>
					<span class="ot-tool__icon" aria-hidden="true" v-html="icons[item.icon ?? 'chat']"></span>
					<span class="ot-tool__arrow" aria-hidden="true" v-html="icons.arrow"></span>
					<h3>{{ item.title }}</h3>
					<p>{{ item.copy }}</p>
				</a>
			</div>
		</section>

		<section class="ot-section ot-limits" data-reveal>
			<div class="ot-bar">
				<h2 class="ot-h2">{{ copy.limits.title }}</h2>
				<p class="ot-bar__note">{{ copy.limits.note }}</p>
			</div>
			<ul>
				<li v-for="item in copy.limits.items" :key="item">{{ item }}</li>
			</ul>
			<a class="ot-btn ot-btn--ghost" :href="withBase(copy.limits.all.href)">{{
				copy.limits.all.text
			}}</a>
		</section>

		<LandingQuickstart />
		<LandingFaq />
		<LandingFinale />
	</div>
</template>

<style scoped>
.ot-grid--two {
	grid-template-columns: repeat(2, minmax(0, 1fr));
}

.ot-copy a {
	color: var(--ot-accent);
	text-decoration: underline;
	text-underline-offset: 3px;
}

.ot-limits ul {
	margin: 0 0 2rem;
	padding-left: 1.2rem;
	list-style: disc;
	max-width: 46rem;
	color: var(--ot-ink-soft);
	line-height: 1.8;
}

@media (max-width: 600px) {
	.ot-grid--two {
		grid-template-columns: minmax(0, 1fr);
	}
}
</style>
