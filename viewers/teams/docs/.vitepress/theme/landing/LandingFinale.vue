<script setup lang="ts">
import { withBase } from 'vitepress';

const columns = [
	{
		title: 'Product',
		links: [
			{ text: 'Live demo', href: '/demo/', external: false },
			{ text: 'Bring your own server', href: '/server', external: false },
			{ text: 'Collaboration', href: '/collaboration', external: false },
			{ text: 'Releases', href: '/releasing', external: false },
		],
	},
	{
		title: 'Docs',
		links: [
			{ text: 'Getting started', href: '/getting-started', external: false },
			{ text: 'Architecture', href: '/architecture', external: false },
			{ text: 'Element API', href: '/api', external: false },
			{ text: 'Limitations', href: '/limitations', external: false },
		],
	},
	{
		title: 'Community',
		links: [
			{ text: 'GitHub', href: 'https://github.com/ChristopherVR/teams-viewer', external: true },
			{
				text: 'Issues',
				href: 'https://github.com/ChristopherVR/teams-viewer/issues',
				external: true,
			},
			{
				text: 'License',
				href: 'https://github.com/ChristopherVR/teams-viewer/blob/main/LICENSE',
				external: true,
			},
		],
	},
];
</script>

<template>
	<section class="ot-section ot-finale">
		<p class="ot-kicker">Get started</p>
		<h2 class="ot-h2">Your workspace, your server.</h2>
		<p class="ot-copy">
			Install a binding, start the reference server with <code>npx openteams-server</code>, and try
			the demo with two tabs. Apache-2.0 licensed, strict TypeScript, and honest about the gaps.
		</p>
		<div class="ot-finale__actions">
			<a class="ot-btn ot-btn--solid" :href="withBase('/getting-started')">Get started</a>
			<a
				class="ot-btn ot-btn--ghost"
				href="https://github.com/ChristopherVR/teams-viewer"
				target="_blank"
				rel="noopener"
				>View on GitHub</a
			>
		</div>
		<div class="ot-finale__cols">
			<nav v-for="col in columns" :key="col.title" :aria-label="col.title">
				<h3>{{ col.title }}</h3>
				<ul>
					<li v-for="link in col.links" :key="link.text">
						<a
							:href="link.external ? link.href : withBase(link.href)"
							:target="link.external ? '_blank' : link.href === '/demo/' ? '_self' : undefined"
							:rel="link.external ? 'noopener' : undefined"
							>{{ link.text }}</a
						>
					</li>
				</ul>
			</nav>
		</div>
		<p class="ot-finale__bottom">
			<span>&copy; 2026 ChristopherVR &middot; Apache-2.0</span>
			<span>OpenTeams &middot; not Microsoft Teams, not affiliated with Microsoft</span>
		</p>
	</section>
</template>

<style scoped>
.ot-copy code {
	font-family: var(--ot-mono);
	font-size: 0.86em;
}

.ot-finale__actions {
	display: flex;
	flex-wrap: wrap;
	gap: 0.8rem;
	margin-top: 1.8rem;
}

.ot-finale__cols {
	display: grid;
	grid-template-columns: repeat(3, minmax(0, 1fr));
	gap: 2rem;
	margin-top: 4rem;
	padding-top: 2.4rem;
	border-top: 1px solid var(--ot-line);
}

.ot-finale__cols h3 {
	margin: 0 0 0.9rem;
	font-family: var(--ot-mono);
	font-size: 0.7rem;
	font-weight: 500;
	letter-spacing: 0.24em;
	text-transform: uppercase;
	color: var(--ot-accent);
}

.ot-finale__cols ul {
	list-style: none;
	margin: 0;
	padding: 0;
	display: grid;
	gap: 0.55rem;
}

.ot-finale__cols a {
	color: var(--ot-ink-soft);
	font-size: 0.92rem;
}

.ot-finale__cols a:hover {
	color: var(--ot-accent);
}

.ot-finale__bottom {
	display: flex;
	flex-wrap: wrap;
	justify-content: space-between;
	gap: 0.6rem 2rem;
	margin: 3rem 0 0;
	font-family: var(--ot-mono);
	font-size: 0.7rem;
	letter-spacing: 0.06em;
	color: var(--ot-ink-soft);
}

@media (max-width: 700px) {
	.ot-finale__cols {
		grid-template-columns: minmax(0, 1fr);
	}
}
</style>
