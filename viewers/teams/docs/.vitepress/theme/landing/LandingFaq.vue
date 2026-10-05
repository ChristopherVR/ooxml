<script setup lang="ts">
import { withBase } from 'vitepress';

const items = [
	{
		q: 'Is it free to use commercially?',
		a: 'The repository is Apache-2.0 licensed. The packages are published to npm as openteams-*-viewer and openteams-server.',
	},
	{
		q: 'Does it need a server?',
		a: 'Real use does: you bring your own server for sync (y-websocket), signaling and optional file storage. The reference server, openteams-server, does all three in one Node process. Without a server the demos run in a local mode that only connects tabs of one browser.',
		href: '/server',
		link: 'Bring your own server',
	},
	{
		q: 'Is it end-to-end encrypted?',
		a: 'No. Chat travels as Yjs updates through your server, which can read and store it. Call media is DTLS-SRTP between peers.',
		href: '/limitations',
		link: 'Limitations',
	},
	{
		q: 'Is this Microsoft Teams?',
		a: 'No. OpenTeams is not Microsoft Teams, is not affiliated with Microsoft and claims no feature parity. It has channels, posts with replies, reactions, files, search, presence and meetings with screen share.',
	},
	{
		q: 'How many people can join a call?',
		a: 'Calls are a full mesh with no media server, so they suit a handful of people (default cap 12), not a webinar. TURN is yours to run.',
		href: '/server#calls-stun-and-turn',
		link: 'Calls and TURN',
	},
	{
		q: 'Which frameworks are supported?',
		a: 'React, Vue 3, Angular, Svelte 5, SolidJS and plain JavaScript. They share one <teams-app> element and one core client; each binding also exposes a raw hook so you can build your own UI.',
		href: '/frameworks/react',
		link: 'Framework guides',
	},
	{
		q: 'Does it work with Word, Excel, PowerPoint and Visio files?',
		a: 'Files are shared as links. Office files are recognised by name so a host application can open them in the matching viewer; OpenTeams does not edit them itself.',
	},
];
</script>

<template>
	<section class="ot-section ot-faq">
		<p class="ot-kicker">FAQ</p>
		<h2 class="ot-h2">Common questions</h2>
		<div class="ot-faq__list">
			<details v-for="item in items" :key="item.q" class="ot-faq__item">
				<summary>{{ item.q }}</summary>
				<p>{{ item.a }}</p>
				<a v-if="item.href" :href="withBase(item.href)">{{ item.link }} &rarr;</a>
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
	font-family: var(--ot-mono);
	font-size: 0.7rem;
	letter-spacing: 0.14em;
	text-transform: uppercase;
	color: var(--ot-accent);
}
</style>
