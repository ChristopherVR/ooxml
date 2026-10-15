<script setup lang="ts">
import { withBase } from 'vitepress';

import CodeCard from './code/CodeCard.vue';
import { CORE_SAMPLE } from './code/samples';
import { useLandingCopy } from './copy';

const copy = useLandingCopy();
</script>

<template>
	<section class="pv-section pv-status">
		<header class="pv-sectionbar">
			<h2 class="pv-h2">{{ copy.status.title }}</h2>
		</header>
		<div class="pv-status__grid">
			<div class="pv-status__copy">
				<p class="pv-copy">{{ copy.status.copy }}</p>
				<a class="pv-link" :href="withBase(copy.status.link.href)">
					{{ copy.status.link.text }}
				</a>
			</div>
			<div class="pv-status__code">
				<CodeCard file="load-save.ts" badge="XLSX &middot; XLS &middot; CSV" :code="CORE_SAMPLE" />
				<aside class="pv-status__note">
					<strong>{{ copy.status.roadmapTitle }}</strong>
					<p>{{ copy.status.roadmapCopy }}</p>
					<a :href="withBase(copy.status.roadmapLink.href)">
						{{ copy.status.roadmapLink.text }}
					</a>
				</aside>
			</div>
		</div>
	</section>
</template>

<style scoped>
.pv-status__grid {
	display: grid;
	grid-template-columns: minmax(0, 1fr) minmax(0, 1.1fr);
	gap: clamp(2rem, 5vw, 5rem);
	align-items: center;
}

.pv-status__copy .pv-copy {
	max-width: 33rem;
}

.pv-status__code {
	display: flex;
	flex-direction: column;
	gap: 1.2rem;
	min-width: 0;
}

.pv-status__note {
	background: var(--pv-surface);
	border: 1px solid var(--pv-line);
	border-left: 3px solid var(--pv-accent);
	border-radius: 6px;
	padding: 1.1rem 1.3rem 1.2rem;
}

.pv-status__note strong {
	font-weight: 600;
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	color: var(--pv-accent);
}

.pv-status__note p {
	margin: 0.6rem 0 0.7rem;
	font-size: 0.88rem;
	line-height: 1.7;
	color: var(--pv-ink-soft);
}

.pv-status__note a {
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	color: var(--pv-accent);
}

@media (max-width: 900px) {
	.pv-status__grid {
		grid-template-columns: minmax(0, 1fr);
	}
}
</style>
