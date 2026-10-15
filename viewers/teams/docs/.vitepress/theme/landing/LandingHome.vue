<script setup lang="ts">
import { withBase } from 'vitepress';

import { icons } from './icons';

import LandingFaq from './LandingFaq.vue';
import LandingFinale from './LandingFinale.vue';
import LandingHero from './LandingHero.vue';
import LandingLiveDemo from './LandingLiveDemo.vue';
import LandingQuickstart from './LandingQuickstart.vue';

const demos = [
	{
		title: 'Vanilla: the whole app',
		copy: 'One <teams-app> element: channels, posts with replies and reactions, files, search, presence and meetings. Open it in two tabs to see two people talk.',
		href: '/demo/',
	},
	{
		title: 'React: component and hook',
		copy: 'The <Teams /> component on the left and a hand-written panel over the useTeams() hook on the right, both on one workspace.',
		href: '/demo-react/',
	},
];

const features = [
	{
		title: 'Channels and chat',
		copy: 'Channels, posts with replies, reactions, author-only edit and delete, search and unread counts, all in one Yjs document that merges concurrent edits.',
		href: '/getting-started',
		link: 'Getting started',
		icon: 'chat',
	},
	{
		title: 'Presence',
		copy: 'Who is online, typing indicators and an availability you set yourself (available, busy, away). Presence is advisory, like any Yjs awareness.',
		href: '/architecture',
		link: 'Architecture',
		icon: 'people',
	},
	{
		title: 'Meetings over WebRTC',
		copy: 'Audio, video, screen share and raise hand in a full-mesh call: no media server, honest up to about a dozen people. TURN is yours to run.',
		href: '/server#calls-stun-and-turn',
		link: 'Calls and TURN',
		icon: 'video',
	},
	{
		title: 'Files in the conversation',
		copy: 'Attachments are uploaded to your server (or your own uploader) and posted as links. Word, Excel, PowerPoint and Visio files are recognised so a host can open them.',
		href: '/server#files',
		link: 'File storage',
		icon: 'file',
	},
	{
		title: 'Bring your own server',
		copy: 'Three small contracts: y-websocket sync, a JSON signaling relay and optional file storage. openteams-server does all three in one Node process.',
		href: '/server',
		link: 'Server guide',
		icon: 'server',
	},
	{
		title: 'Six frameworks',
		copy: 'React, Vue, Angular, Svelte, Solid and vanilla bindings. Each gives you the full component and a raw hook over the same client store.',
		href: '/frameworks/react',
		link: 'Framework guides',
		icon: 'code',
	},
];

const limits = [
	'No end-to-end encryption: your server stores and can read every message and file.',
	'Authorization is the server’s job; the reference server has one shared token, not user accounts.',
	'Mesh calls suit a handful of people, not a webinar.',
	'No 1:1 chats, notifications, calendar, recording, background blur or live captions.',
	'Not Microsoft Teams, not affiliated with Microsoft, and no claim of feature parity.',
];
</script>

<template>
	<div class="ot-landing">
		<LandingHero />

		<LandingLiveDemo />

		<section class="ot-section">
			<div class="ot-bar">
				<h2 class="ot-h2">Two ways to embed it</h2>
				<p class="ot-bar__note">Demo apps</p>
			</div>
			<div class="ot-grid ot-grid--two">
				<article v-for="demo in demos" :key="demo.href" class="ot-card">
					<h3>{{ demo.title }}</h3>
					<p>{{ demo.copy }}</p>
					<a class="ot-more" :href="withBase(demo.href)" target="_self">Open the demo</a>
				</article>
			</div>
		</section>

		<section class="ot-section">
			<div class="ot-bar">
				<h2 class="ot-h2">A workspace you can host</h2>
				<p class="ot-bar__note">What it does</p>
			</div>
			<div class="ot-grid">
				<a
					v-for="item in features"
					:key="item.title"
					class="ot-card ot-tool"
					:href="withBase(item.href)"
					:aria-label="`${item.title}: ${item.link}`"
				>
					<span class="ot-tool__icon" aria-hidden="true" v-html="icons[item.icon]"></span>
					<span class="ot-tool__arrow" aria-hidden="true" v-html="icons.arrow"></span>
					<h3>{{ item.title }}</h3>
					<p>{{ item.copy }}</p>
				</a>
			</div>
		</section>

		<section class="ot-section ot-limits">
			<div class="ot-bar">
				<h2 class="ot-h2">What it is not</h2>
				<p class="ot-bar__note">Honest limits</p>
			</div>
			<ul>
				<li v-for="item in limits" :key="item">{{ item }}</li>
			</ul>
			<a class="ot-btn ot-btn--ghost" :href="withBase('/limitations')">All limitations</a>
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
