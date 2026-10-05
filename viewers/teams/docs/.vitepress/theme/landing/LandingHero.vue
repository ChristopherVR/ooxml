<script setup lang="ts">
import { withBase } from 'vitepress';
</script>

<template>
	<section class="ot-hero">
		<div class="ot-hero__inner">
			<div class="ot-hero__copy">
				<p class="ot-kicker ot-enter">Open source &middot; bring your own server</p>
				<h1 class="ot-hero__title ot-enter ot-enter--2">
					Team chat and calls,<br />
					<em>on a server you run.</em>
				</h1>
				<p class="ot-hero__sub ot-enter ot-enter--3">
					OpenTeams is a team workspace for the web: channels and chat, presence, WebRTC meetings
					and Office files shared in the conversation. One web component, thin bindings for six
					frameworks, and a small reference server. An early implementation: not Microsoft Teams,
					and no end-to-end encryption.
				</p>
				<div class="ot-hero__actions ot-enter ot-enter--3">
					<a class="ot-btn ot-btn--solid" :href="withBase('/getting-started')">Get started</a>
					<a class="ot-btn ot-btn--ghost" :href="withBase('/demo/')" target="_self">Try the demo</a>
				</div>
				<pre class="ot-hero__install ot-enter ot-enter--3"><code>npm install openteams-react-viewer
npx openteams-server</code></pre>
			</div>
			<div class="ot-hero__stage ot-enter ot-enter--2">
				<a class="ot-hero__frame" :href="withBase('/demo/')" target="_self">
					<span class="ot-hero__framebar">
						<span class="ot-hero__dots" aria-hidden="true"><i></i><i></i><i></i></span>
						<span>&lt;teams-app&gt; in local mode</span>
						<span class="ot-hero__try">Open the demo &rarr;</span>
					</span>
					<img
						:src="withBase('/hero-workspace.png')"
						alt="The OpenTeams demo: a channel list, a conversation with replies and reactions, and the composer"
						width="1280"
						height="800"
						fetchpriority="high"
					/>
				</a>
			</div>
		</div>
	</section>
</template>

<style scoped>
.ot-hero {
	background-image: radial-gradient(
		color-mix(in srgb, var(--ot-ink) 9%, transparent) 1px,
		transparent 1px
	);
	background-size: 26px 26px;
}

.ot-hero__inner {
	max-width: 84rem;
	margin: 0 auto;
	padding: clamp(2.5rem, 8vh, 6rem) clamp(1rem, 5vw, 4.5rem) clamp(3rem, 9vh, 6rem);
	display: grid;
	grid-template-columns: minmax(0, 1fr) minmax(0, 1.05fr);
	gap: clamp(2rem, 5vw, 4.5rem);
	align-items: center;
}

.ot-hero__copy {
	min-width: 0;
}

.ot-hero__title {
	font-family: var(--ot-display);
	font-size: clamp(2.5rem, 5.6vw, 4.6rem);
	font-weight: 680;
	line-height: 1.02;
	letter-spacing: -0.025em;
	margin: 1.3rem 0 1.4rem;
}

.ot-hero__title em {
	font-style: italic;
	font-weight: 460;
	color: var(--ot-accent);
}

.ot-hero__sub {
	font-size: clamp(1rem, 1.4vw, 1.1rem);
	line-height: 1.75;
	color: var(--ot-ink-soft);
	max-width: 34rem;
	margin: 0;
}

.ot-hero__actions {
	display: flex;
	flex-wrap: wrap;
	gap: 0.9rem;
	margin-top: 2rem;
}

.ot-hero__install {
	margin: 1.6rem 0 0;
	padding: 0.9rem 1.1rem;
	border-radius: 6px;
	background: #14171b;
	color: #e6e4df;
	font-family: var(--ot-mono);
	font-size: 0.82rem;
	line-height: 1.7;
	overflow-x: auto;
	max-width: 34rem;
}

.ot-hero__frame {
	display: block;
	background: var(--ot-surface);
	border: 1px solid var(--ot-line);
	border-radius: 8px;
	overflow: hidden;
	box-shadow: var(--ot-shadow);
	transition:
		transform 0.35s cubic-bezier(0.25, 0.6, 0.3, 1),
		border-color 0.35s ease;
}

.ot-hero__frame:hover {
	transform: translateY(-4px);
	border-color: var(--ot-accent);
}

.ot-hero__framebar {
	display: flex;
	align-items: center;
	gap: 0.9rem;
	padding: 0.6rem 0.9rem;
	border-bottom: 1px solid var(--ot-line);
	font-family: var(--ot-mono);
	font-size: 0.68rem;
	color: var(--ot-ink-soft);
}

.ot-hero__dots {
	display: inline-flex;
	gap: 0.3rem;
}

.ot-hero__dots i {
	width: 8px;
	height: 8px;
	border-radius: 50%;
	background: var(--ot-line);
}

.ot-hero__try {
	margin-left: auto;
	color: var(--ot-accent);
	letter-spacing: 0.14em;
	text-transform: uppercase;
	white-space: nowrap;
}

.ot-hero__frame img {
	display: block;
	width: 100%;
	height: auto;
}

@media (max-width: 900px) {
	.ot-hero__inner {
		grid-template-columns: minmax(0, 1fr);
	}
}
</style>
