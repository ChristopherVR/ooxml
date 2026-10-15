<script setup lang="ts">
import { ref } from 'vue';

import { useLandingCopy } from './copy';
import LiveDemoPane from './LiveDemoPane.vue';
import { DEMO_FRAMEWORKS, useLiveDemo } from './useLiveDemo';

/**
 * Live, embedded demo section: the demo apps built from this repository run
 * inside the landing page with a framework switcher, opening the sample workbook.
 */
const copy = useLandingCopy();
const section = ref<HTMLElement | null>(null);
const { started, activeKey, src, activeLabel, start, selectFramework } = useLiveDemo(section);
</script>

<template>
	<section id="live-demo" ref="section" class="pv-section pv-live">
		<header class="pv-sectionbar">
			<h2 class="pv-h2">{{ copy.demos.title }}</h2>
			<p class="pv-sectionbar__note">{{ copy.demos.copy }}</p>
		</header>

		<div class="pv-live__controls">
			<div class="pv-live__tabs" role="tablist" :aria-label="copy.demos.frameworkLabel">
				<button
					v-for="f in DEMO_FRAMEWORKS"
					:key="f.key"
					type="button"
					role="tab"
					class="pv-live__tab"
					:class="{ 'is-active': f.key === activeKey }"
					:aria-selected="f.key === activeKey"
					@click="selectFramework(f.key)"
				>
					{{ f.label }}
				</button>
			</div>
		</div>

		<div v-if="!started" class="pv-live__poster">
			<button type="button" class="pv-btn pv-btn--solid" @click="start">
				<span>{{ copy.demos.load }}</span>
			</button>
		</div>
		<template v-else>
			<div class="pv-live__stage">
				<LiveDemoPane
					:key="src"
					:src="src"
					:title="`${activeLabel} · xlsx-viewer ${copy.demos.kicker}`"
					:caption="`${activeLabel} · Sample workbook.xlsx`"
					:open-label="copy.demos.openFull"
					:loading-label="copy.demos.loading"
				/>
			</div>
			<p class="pv-live__hint">{{ copy.demos.hint }}</p>
		</template>
	</section>
</template>

<style scoped>
.pv-live__controls {
	display: flex;
	flex-wrap: wrap;
	align-items: center;
	gap: 0.9rem 1.4rem;
	margin-top: 2.2rem;
}

.pv-live__tabs {
	display: inline-flex;
	flex-wrap: wrap;
	gap: 0.35rem;
	padding: 0.3rem;
	background: var(--pv-surface);
	border: 1px solid var(--pv-line);
	border-radius: 6px;
}

.pv-live__tab {
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	font-weight: 500;
	color: var(--pv-ink-soft);
	padding: 0.5rem 0.9rem;
	border-radius: 4px;
	transition:
		color 0.18s ease,
		background-color 0.18s ease;
}

.pv-live__tab:hover {
	color: var(--pv-ink);
}

.pv-live__tab.is-active {
	background: var(--pv-accent-soft);
	color: var(--pv-accent);
}

.pv-live__poster {
	display: flex;
	align-items: center;
	justify-content: center;
	min-height: 340px;
	margin-top: 1.4rem;
	border: 1px dashed var(--pv-line);
	border-radius: 8px;
	background: var(--pv-surface);
}

.pv-live__stage {
	display: grid;
	grid-template-columns: minmax(0, 1fr);
	height: clamp(480px, 72vh, 760px);
	margin-top: 1.4rem;
}

.pv-live__hint {
	margin-top: 1.1rem;
	font-size: 0.86rem;
	line-height: 1.65;
	color: var(--pv-ink-soft);
	max-width: 46rem;
}
</style>
