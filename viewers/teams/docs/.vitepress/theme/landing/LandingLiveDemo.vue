<script setup lang="ts">
import { withBase } from 'vitepress';
import { computed, onBeforeUnmount, onMounted, ref } from 'vue';

/**
 * Embedded live demo. The demo apps are built next to the docs and run in the
 * core's local mode: chat, presence and calls travel between tabs and frames of
 * the visitor's own browser over BroadcastChannel, so the two-person mode puts
 * two same-origin frames in one room and they really talk to each other. There
 * is no server behind them and nothing leaves the browser.
 */
const demos = [
	{ key: 'vanilla', label: 'Vanilla', path: '/demo/' },
	{ key: 'react', label: 'React', path: '/demo-react/' },
	{ key: 'vue', label: 'Vue', path: '/demo-vue/' },
	{ key: 'angular', label: 'Angular', path: '/demo-angular/' },
	{ key: 'svelte', label: 'Svelte', path: '/demo-svelte/' },
	{ key: 'solid', label: 'Solid', path: '/demo-solid/' },
];

const section = ref<HTMLElement | null>(null);
const started = ref(false);
const mode = ref<'solo' | 'duo'>('solo');
const activeKey = ref('vanilla');
/** Bob's app. Each pane picks its own framework; both join the same room. */
const guestKey = ref('react');
const room = ref(randomRoom());

function randomRoom(): string {
	return `landing-${Math.random().toString(36).slice(2, 10)}`;
}

const active = computed(() => demos.find((d) => d.key === activeKey.value) ?? demos[0]);
const base = computed(() => withBase(active.value.path));
const soloSrc = computed(() => `${base.value}?name=Ada&room=${room.value}`);
const guest = computed(() => demos.find((d) => d.key === guestKey.value) ?? demos[1]);
const adaSrc = computed(() => `${base.value}?name=Ada&room=${room.value}`);
const bobSrc = computed(() => `${withBase(guest.value.path)}?name=Bob&room=${room.value}`);

function select(key: string): void {
	activeKey.value = key;
	if (mode.value === 'solo') room.value = randomRoom();
}

function selectGuest(key: string): void {
	guestKey.value = key;
}

function setMode(next: 'solo' | 'duo'): void {
	mode.value = next;
	room.value = randomRoom();
	started.value = true;
}

let observer: IntersectionObserver | null = null;
onMounted(() => {
	if (!section.value || typeof IntersectionObserver === 'undefined') return;
	observer = new IntersectionObserver(
		(entries) => {
			if (entries.some((entry) => entry.isIntersecting)) {
				started.value = true;
				observer?.disconnect();
				observer = null;
			}
		},
		{ rootMargin: '300px 0px' },
	);
	observer.observe(section.value);
});
onBeforeUnmount(() => observer?.disconnect());
</script>

<template>
	<section id="live-demo" ref="section" class="ot-section ot-live">
		<p class="ot-kicker">Live demo</p>
		<h2 class="ot-h2">Try it right here</h2>
		<p class="ot-copy">
			This is the real app running in your browser, with no server behind it: chat, presence and
			calls travel only between frames and tabs of this browser (BroadcastChannel). Switch to two
			people to put Ada and Bob in one room, each in the framework you choose (for example React for
			Ada and Vue for Bob), and watch a message cross from one pane to the other. Real use needs a
			server; see <a :href="withBase('/server')">bring your own server</a>.
		</p>

		<div class="ot-live__controls">
			<div class="ot-live__tabs" role="tablist" aria-label="Ada's framework">
				<button
					v-for="d in demos"
					:key="d.key"
					type="button"
					role="tab"
					class="ot-live__tab"
					:class="{ 'is-active': d.key === activeKey }"
					:aria-selected="d.key === activeKey"
					@click="select(d.key)"
				>
					{{ d.label }}
				</button>
			</div>
			<div class="ot-live__tabs">
				<button
					type="button"
					class="ot-live__tab"
					:class="{ 'is-active': mode === 'solo' }"
					@click="setMode('solo')"
				>
					One person
				</button>
				<button
					type="button"
					class="ot-live__tab"
					:class="{ 'is-active': mode === 'duo' }"
					@click="setMode('duo')"
				>
					Two people
				</button>
			</div>
			<label v-if="mode === 'duo'" class="ot-live__guestpick">
				<span>Bob's framework</span>
				<select :value="guestKey" @change="selectGuest(($event.target as HTMLSelectElement).value)">
					<option v-for="d in demos" :key="d.key" :value="d.key">{{ d.label }}</option>
				</select>
			</label>
		</div>

		<div v-if="!started" class="ot-live__poster">
			<button type="button" class="ot-btn ot-btn--solid" @click="started = true">
				Load the live demo
			</button>
		</div>
		<div v-else-if="mode === 'solo'" class="ot-live__stage">
			<figure class="ot-live__pane">
				<figcaption>
					<span>{{ active.label }} &middot; Ada</span>
					<a :href="soloSrc" target="_blank" rel="noopener">Open full app &rarr;</a>
				</figcaption>
				<iframe
					:key="soloSrc"
					:src="soloSrc"
					:title="`${active.label} OpenTeams demo`"
					allow="clipboard-read; clipboard-write; fullscreen; camera; microphone; display-capture"
				></iframe>
			</figure>
		</div>
		<div v-else class="ot-live__stage ot-live__stage--duo">
			<figure class="ot-live__pane">
				<figcaption>
					<span>{{ active.label }} &middot; Ada</span>
					<a :href="adaSrc" target="_blank" rel="noopener">Open full app &rarr;</a>
				</figcaption>
				<iframe
					:key="adaSrc"
					:src="adaSrc"
					title="Ada"
					allow="clipboard-read; clipboard-write"
				></iframe>
			</figure>
			<figure class="ot-live__pane">
				<figcaption>
					<span>{{ guest.label }} &middot; Bob</span>
					<a :href="bobSrc" target="_blank" rel="noopener">Open full app &rarr;</a>
				</figcaption>
				<iframe
					:key="bobSrc"
					:src="bobSrc"
					title="Bob"
					allow="clipboard-read; clipboard-write"
				></iframe>
			</figure>
		</div>
		<p v-if="started" class="ot-live__hint">
			{{
				mode === 'duo'
					? 'Two separate apps, each built from its own framework binding, join one room by its name and share it through BroadcastChannel. The framework does not matter: they speak the same core protocol. This is the local mode of the core, not a hosted service: with your own server, the same two people can be on different machines.'
					: 'Everything runs client-side and stays in this browser. Open the full app, or switch to two people to see presence and chat between two users.'
			}}
		</p>
	</section>
</template>

<style scoped>
.ot-copy :deep(a),
.ot-copy a {
	color: var(--ot-accent);
	text-decoration: underline;
	text-underline-offset: 3px;
}

.ot-live__controls {
	display: flex;
	flex-wrap: wrap;
	align-items: center;
	gap: 0.9rem 1.4rem;
	margin-top: 2.2rem;
}

.ot-live__tabs {
	display: inline-flex;
	flex-wrap: wrap;
	gap: 0.35rem;
	padding: 0.3rem;
	background: var(--ot-surface);
	border: 1px solid var(--ot-line);
	border-radius: 6px;
}

.ot-live__tab {
	font-family: var(--ot-mono);
	font-size: 0.72rem;
	font-weight: 500;
	letter-spacing: 0.08em;
	text-transform: uppercase;
	color: var(--ot-ink-soft);
	background: transparent;
	border: 0;
	border-radius: 4px;
	padding: 0.5rem 0.85rem;
	cursor: pointer;
}

.ot-live__tab:hover {
	color: var(--ot-ink);
}

.ot-live__tab.is-active {
	background: var(--ot-accent);
	color: #fff;
}

.ot-live__poster {
	display: grid;
	place-items: center;
	min-height: 18rem;
	margin-top: 1.4rem;
	background: var(--ot-surface);
	border: 1px solid var(--ot-line);
	border-radius: 8px;
}

.ot-live__stage {
	margin-top: 1.4rem;
}

.ot-live__stage--duo {
	display: grid;
	grid-template-columns: repeat(2, minmax(0, 1fr));
	gap: 1rem;
}

.ot-live__pane {
	margin: 0;
	display: flex;
	flex-direction: column;
	border: 1px solid var(--ot-line);
	border-radius: 8px;
	overflow: hidden;
	background: var(--ot-surface);
}

.ot-live__pane figcaption {
	display: flex;
	justify-content: space-between;
	gap: 1rem;
	padding: 0.55rem 0.9rem;
	font-family: var(--ot-mono);
	font-size: 0.7rem;
	letter-spacing: 0.08em;
	text-transform: uppercase;
	border-bottom: 1px solid var(--ot-line);
	color: var(--ot-ink-soft);
}

.ot-live__pane figcaption a {
	color: var(--ot-accent);
}

.ot-live__pane iframe {
	width: 100%;
	height: 640px;
	border: 0;
	background: var(--ot-surface);
}

.ot-live__guestpick {
	display: inline-flex;
	align-items: center;
	gap: 0.6rem;
	font-family: var(--ot-mono);
	font-size: 0.72rem;
	letter-spacing: 0.08em;
	text-transform: uppercase;
	color: var(--ot-ink-soft);
}

.ot-live__guestpick select {
	font: inherit;
	color: var(--ot-ink);
	background: var(--ot-surface);
	border: 1px solid var(--ot-line);
	border-radius: 4px;
	padding: 0.45rem 0.6rem;
}

.ot-live__hint {
	margin: 1rem 0 0;
	max-width: 48rem;
	font-size: 0.85rem;
	line-height: 1.7;
	color: var(--ot-ink-soft);
}

@media (max-width: 900px) {
	.ot-live__stage--duo {
		grid-template-columns: minmax(0, 1fr);
	}
}
</style>
