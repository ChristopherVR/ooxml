<script setup lang="ts">
import { withBase } from 'vitepress';
import { ref } from 'vue';

/** Snippets mirror docs/frameworks/*.md; keep them in sync with those pages. */
const frameworks = [
	{
		id: 'react',
		label: 'React',
		install: 'npm install openteams-react-viewer react',
		href: '/frameworks/react',
		code: `import { Teams } from 'openteams-react-viewer';

export function App() {
	return (
		<Teams
			workspaceId="acme"
			userName="Ada"
			config={{ mode: 'local', iceServers: [] }}
		/>
	);
}`,
	},
	{
		id: 'vue',
		label: 'Vue',
		install: 'npm install openteams-vue-viewer vue',
		href: '/frameworks/vue',
		code: `<script setup lang="ts">
import { Teams } from 'openteams-vue-viewer';
const config = { mode: 'local', iceServers: [] } as const;
<\/script>

<template>
	<Teams workspace-id="acme" user-name="Ada" :config="config" />
</template>`,
	},
	{
		id: 'vanilla',
		label: 'Vanilla JS',
		install: 'npm install openteams-vanilla-viewer',
		href: '/frameworks/vanilla',
		code: `import { mountTeams } from 'openteams-vanilla-viewer';

const teams = mountTeams(document.getElementById('teams')!, {
	workspaceId: 'acme',
	userName: 'Ada',
	config: { mode: 'local', iceServers: [] },
});

teams.destroy(); // when you are done`,
	},
];

const active = ref(frameworks[0]);
</script>

<template>
	<section class="ot-section ot-quick">
		<p class="ot-kicker">Quickstart</p>
		<h2 class="ot-h2">A workspace in a few lines</h2>
		<p class="ot-copy">
			Install the binding for your framework and render the component. Local mode needs no server;
			for real use point <code>config</code> at your own sync and signaling endpoints (the reference
			server does both). Angular, Svelte and Solid have the same options.
		</p>
		<div class="ot-quick__panel">
			<div class="ot-quick__tabs" role="tablist">
				<button
					v-for="fw in frameworks"
					:key="fw.id"
					type="button"
					role="tab"
					:aria-selected="fw.id === active.id"
					:class="['ot-quick__tab', { 'is-active': fw.id === active.id }]"
					@click="active = fw"
				>
					{{ fw.label }}
				</button>
			</div>
			<pre class="ot-quick__install"><code>{{ active.install }}</code></pre>
			<pre class="ot-quick__code"><code>{{ active.code }}</code></pre>
			<a class="ot-quick__link" :href="withBase(active.href)"
				>Full guide: {{ active.label }} &rarr;</a
			>
		</div>
	</section>
</template>

<style scoped>
.ot-copy code {
	font-family: var(--ot-mono);
	font-size: 0.86em;
}

.ot-quick__panel {
	margin-top: 2rem;
	max-width: 52rem;
}

.ot-quick__tabs {
	display: flex;
	gap: 0.35rem;
	margin-bottom: 0.8rem;
}

.ot-quick__tab {
	font-family: var(--ot-mono);
	font-size: 0.72rem;
	letter-spacing: 0.08em;
	text-transform: uppercase;
	color: var(--ot-ink-soft);
	background: var(--ot-surface);
	border: 1px solid var(--ot-line);
	border-radius: 4px;
	padding: 0.5rem 0.85rem;
	cursor: pointer;
}

.ot-quick__tab.is-active {
	background: var(--ot-accent);
	border-color: var(--ot-accent);
	color: #fff;
}

.ot-quick__install,
.ot-quick__code {
	margin: 0 0 0.8rem;
	padding: 1rem 1.2rem;
	background: #14171b;
	color: #f0efec;
	border: 1px solid var(--ot-line);
	border-radius: 8px;
	overflow-x: auto;
	font-family: var(--ot-mono);
	font-size: 0.82rem;
	line-height: 1.65;
}

.ot-quick__link {
	font-family: var(--ot-mono);
	font-size: 0.72rem;
	letter-spacing: 0.14em;
	text-transform: uppercase;
	color: var(--ot-accent);
}
</style>
