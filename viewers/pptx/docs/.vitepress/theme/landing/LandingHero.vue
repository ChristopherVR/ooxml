<script setup lang="ts">
import { withBase } from 'vitepress';

import InstallPicker from './code/InstallPicker.vue';
import { useLandingCopy } from './copy';
import LandingRings from './LandingRings.vue';

const copy = useLandingCopy();
</script>

<template>
	<section class="pv-hero">
		<div class="pv-hero__inner">
			<div class="pv-hero__copy">
				<h1 class="pv-hero__title pv-enter pv-enter--2">
					{{ copy.hero.titleTop }}<br />
					<em>{{ copy.hero.titleAccent }}</em>
				</h1>
				<p class="pv-hero__sub pv-enter pv-enter--3">{{ copy.hero.sub }}</p>
				<div class="pv-hero__actions pv-enter pv-enter--4">
					<a class="pv-btn pv-btn--solid" :href="withBase(copy.hero.start.href)">
						<span>{{ copy.hero.start.text }}</span>
					</a>
					<a class="pv-btn pv-btn--ghost" href="#live-demo">
						<span>{{ copy.hero.demo }}</span>
					</a>
				</div>
				<div class="pv-enter pv-enter--4">
					<InstallPicker />
				</div>
				<ul class="pv-hero__facts pv-enter pv-enter--4">
					<li v-for="fact in copy.hero.kicker.split(' · ')" :key="fact">
						{{ fact }}
					</li>
				</ul>
			</div>
			<div class="pv-hero__stage pv-enter pv-enter--3">
				<LandingRings />
				<a class="pv-hero__frame" href="#live-demo">
					<span class="pv-hero__framebar">
						<span class="pv-hero__badge" aria-hidden="true">P</span>
						<span class="pv-hero__framefile">{{ copy.hero.frameCaption }}</span>
						<span class="pv-hero__frametry">{{ copy.hero.frameTry }}</span>
					</span>
					<img
						:src="withBase('/hero-viewer.gif')"
						:alt="copy.hero.frameAlt"
						width="960"
						height="540"
						fetchpriority="high"
					/>
				</a>
			</div>
		</div>
		<div class="pv-hero__cue" aria-hidden="true">
			<span>{{ copy.hero.scroll }}</span>
			<i></i>
		</div>
	</section>
</template>

<style scoped>
.pv-hero {
	position: relative;
	min-height: calc(100vh - var(--vp-nav-height, 64px));
	display: flex;
	align-items: center;
}

.pv-hero__inner {
	width: 100%;
	max-width: 84rem;
	margin: 0 auto;
	padding: clamp(3rem, 8vh, 6rem) clamp(1.4rem, 5vw, 4.5rem) clamp(5rem, 12vh, 8rem);
	display: grid;
	grid-template-columns: minmax(0, 1.05fr) minmax(0, 1fr);
	gap: clamp(2rem, 5vw, 5rem);
	align-items: center;
}

.pv-hero__title {
	font-family: var(--pv-display);
	font-size: clamp(2.7rem, 6.2vw, 5.2rem);
	font-weight: 680;
	line-height: 0.99;
	letter-spacing: -0.025em;
	margin: 1.4rem 0 1.6rem;
}

.pv-hero__title em {
	font-style: italic;
	font-weight: 460;
	color: var(--pv-accent);
}

.pv-hero__title:lang(zh-CN) {
	font-size: clamp(2.2rem, 4.8vw, 4rem);
	line-height: 1.25;
	letter-spacing: 0;
}

.pv-hero__title:lang(zh-CN) em {
	font-style: normal;
}

.pv-hero__sub {
	font-size: clamp(1rem, 1.4vw, 1.12rem);
	line-height: 1.75;
	color: var(--pv-ink-soft);
	max-width: 33rem;
}

.pv-hero__actions {
	display: flex;
	flex-wrap: wrap;
	gap: 0.9rem;
	margin-top: 2.2rem;
}

.pv-hero__frame {
	display: block;
	background: var(--pv-surface);
	border: 1px solid var(--pv-line);
	border-radius: 8px;
	overflow: hidden;
	transition:
		border-color 200ms var(--pv-ease),
		transform 200ms var(--pv-ease);
}

.pv-hero__frame:hover {
	border-color: var(--pv-accent);
}

@media (prefers-reduced-motion: no-preference) {
	.pv-hero__frame:hover {
		transform: translateY(-2px);
	}
}

.pv-hero__framebar {
	display: flex;
	align-items: center;
	gap: 0.9rem;
	padding: 0.6rem 0.9rem;
	border-bottom: 1px solid var(--pv-line);
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	color: var(--pv-ink-soft);
}

.pv-hero__badge {
	display: inline-grid;
	place-items: center;
	flex: none;
	width: 18px;
	height: 18px;
	border-radius: 3px;
	background: #c43e1c;
	box-shadow: inset -3px 0 #ffffff25;
	color: #fff;
	font-family: var(--pv-ui);
	font-size: 0.66rem;
	font-weight: 700;
	line-height: 1;
}

/* The facts under the hero: status, licence and language, as tags. */
.pv-hero__facts {
	display: flex;
	flex-wrap: wrap;
	gap: 0.4rem;
	margin: 1.2rem 0 0;
	padding: 0;
	list-style: none;
}

.pv-hero__facts li {
	padding: 0.3rem 0.55rem;
	border: 1px solid var(--pv-line);
	border-radius: 4px;
	background: var(--pv-surface);
	color: var(--pv-ink-soft);
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	line-height: 1.2;
}

.pv-hero__facts [data-tone] {
	border-color: color-mix(in srgb, var(--tone) 40%, transparent);
	color: var(--tone);
	font-weight: 600;
}

.pv-hero__facts [data-tone='stable'] {
	--tone: #12715a;
}

.dark .pv-hero__facts [data-tone='stable'] {
	--tone: #4fc08a;
}

.pv-hero__facts [data-tone='beta'] {
	--tone: var(--pv-accent);
}

.pv-hero__frametry {
	margin-left: auto;
	color: var(--pv-accent);
	white-space: nowrap;
}

.pv-hero__frame img {
	display: block;
	width: 100%;
	height: auto;
}

.pv-hero__cue {
	position: absolute;
	bottom: 1.6rem;
	left: 50%;
	transform: translateX(-50%);
	display: flex;
	flex-direction: column;
	align-items: center;
	gap: 0.6rem;
}

.pv-hero__cue span {
	font-family: var(--pv-ui);
	font-size: 0.75rem;
	color: var(--pv-ink-soft);
}

.pv-hero__cue i {
	width: 1px;
	height: 44px;
	background: var(--pv-line);
	overflow: hidden;
	position: relative;
}

.pv-hero__cue i::after {
	content: '';
	position: absolute;
	inset: 0;
	background: var(--pv-accent);
	animation: pv-cue 2s cubic-bezier(0.7, 0, 0.3, 1) infinite;
}

@keyframes pv-cue {
	0% {
		transform: translateY(-100%);
	}

	55%,
	100% {
		transform: translateY(100%);
	}
}

@media (max-width: 900px) {
	.pv-hero {
		min-height: 0;
	}

	.pv-hero__inner {
		grid-template-columns: 1fr;
	}

	.pv-hero__cue {
		display: none;
	}
}

@media (prefers-reduced-motion: reduce) {
	.pv-hero__cue i::after {
		animation: none;
	}
}
</style>
