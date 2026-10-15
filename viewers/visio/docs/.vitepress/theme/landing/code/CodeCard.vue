<script setup lang="ts">
import { computed } from 'vue';

import { highlight } from './highlight';

const props = defineProps<{
	file: string;
	badge?: string;
	code: string;
}>();

const html = computed(() => highlight(props.code));
</script>

<template>
	<div class="pv-code">
		<div class="pv-code__bar">
			<span class="pv-code__dots" aria-hidden="true"><i></i><i></i><i></i></span>
			<span class="pv-code__file">{{ file }}</span>
			<span v-if="badge" class="pv-code__badge">{{ badge }}</span>
		</div>
		<pre><code v-html="html"></code></pre>
	</div>
</template>

<style scoped>
.pv-code {
	background: #181a1e;
	border: 1px solid #34383f;
	border-radius: 8px;
	overflow: hidden;
}

.dark .pv-code {
	background: #202228;
}

.pv-code__bar {
	display: flex;
	align-items: center;
	gap: 0.9rem;
	padding: 0.65rem 1.1rem;
	border-bottom: 1px solid rgba(240, 240, 240, 0.1);
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	color: rgba(240, 240, 240, 0.55);
}

.pv-code__dots {
	display: inline-flex;
	gap: 0.3rem;
}

.pv-code__dots i {
	width: 8px;
	height: 8px;
	border-radius: 50%;
	background: rgba(240, 240, 240, 0.18);
}

.pv-code__badge {
	margin-left: auto;
	color: #7b93dc;
	white-space: nowrap;
}

.pv-code pre {
	margin: 0;
	padding: 1.3rem 1.4rem;
	overflow-x: auto;
}

.pv-code code {
	font-family: var(--pv-mono);
	font-size: 0.78rem;
	line-height: 1.75;
	color: #dfe3e8;
	background: none;
	white-space: pre;
}

.pv-code :deep(.k) {
	color: #7b93dc;
}

.pv-code :deep(.s) {
	color: #a8c795;
}

.pv-code :deep(.t) {
	color: #8fb8d8;
}

.pv-code :deep(.c) {
	color: rgba(223, 227, 232, 0.4);
	font-style: italic;
}
</style>
