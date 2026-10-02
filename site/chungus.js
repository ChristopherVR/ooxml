import { APPS, appIcon, isLive } from './apps.js';
import { disintegrate } from './chungus-dust.js';
import { createSound } from './chungus-sound.js';

/**
 * The Chungus button: a very large rabbit walks in, snaps a gold gauntlet and
 * turns the subscription suite to dust, leaving the open-source apps behind.
 * Pure decoration; it holds no Office logic and never touches the apps. The
 * 3D scene (three.js from a CDN) loads only when the button is pressed.
 */

/** The suite being snapped. Plain letter tiles drawn here, not vendor logos. */
const OLD_SUITE = [
	{ name: 'PowerPoint', letter: 'P', color: '#b7472a' },
	{ name: 'Word', letter: 'W', color: '#2b579a' },
	{ name: 'Excel', letter: 'X', color: '#217346' },
	{ name: 'Visio', letter: 'V', color: '#3955a3' },
];

const LOCK = `<svg viewBox="0 0 16 16" aria-hidden="true"><rect x="3" y="7" width="10" height="7" rx="1.5" fill="currentColor"/><path d="M5 7V5a3 3 0 0 1 6 0v2" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>`;

/** How often the button turns up on a page load; `?chungus` always shows it. */
const APPEARANCE_CHANCE = 0.2;

const ABORT = Symbol('abort');
const reducedMotion = () => matchMedia('(prefers-reduced-motion: reduce)').matches;
const escape = (text) => text.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);

function oldCards(apps, offset) {
	return apps
		.map(
			(app, i) => `<li class="chungus__card" style="--c:${app.color};--i:${i + offset}">
				<span class="chungus__letter">${app.letter}</span>
				<span class="chungus__name">${escape(app.name)}</span>
				<span class="chungus__lock">${LOCK}</span>
			</li>`,
		)
		.join('');
}

function newCards(apps, offset) {
	return apps
		.map((app, i) => {
			const tag = app.tag
				? `<span class="tag tag--${app.tag.tone}">${escape(app.tag.label)}</span>`
				: '<span class="tag tag--live">Live</span>';
			const body = `${appIcon(app)}<span class="chungus__app-name">${escape(app.name)}</span>${tag}`;
			const style = `style="--app:${app.color};--i:${i + offset}"`;
			return isLive(app)
				? `<li><a class="chungus__app" href="#/${app.id}" ${style}>${body}</a></li>`
				: `<li><span class="chungus__app chungus__app--off" ${style}>${body}</span></li>`;
		})
		.join('');
}

function side(name, oldApps, oldOffset, newApps, newOffset) {
	return `<div class="chungus__side chungus__side--${name}">
		<ul class="chungus__list chungus__list--old">${oldCards(oldApps, oldOffset)}</ul>
		<ul class="chungus__list chungus__list--new">${newCards(newApps, newOffset)}</ul>
	</div>`;
}

function markup() {
	return `<div class="chungus__world"></div>
		<div class="chungus__stage">
			<header class="chungus__head">
				<p class="chungus__eyebrow" data-old>Per user, per month, forever</p>
				<p class="chungus__eyebrow" data-new>Apache-2.0, free for everyone</p>
				<h2 class="chungus__title" data-old>The subscription suite</h2>
				<h2 class="chungus__title" data-new>OOXML Office</h2>
			</header>
			${side('left', OLD_SUITE.slice(0, 2), 0, APPS.slice(0, 2), 0)}
			${side('right', OLD_SUITE.slice(2), 2, APPS.slice(2), 2)}
			<footer class="chungus__foot">
				<p class="chungus__caption" aria-live="polite"></p>
				<div class="chungus__actions">
					<button class="button button--page" type="button" data-close>Back to the suite</button>
				</div>
			</footer>
		</div>
		<p class="chungus__bubble" aria-hidden="true">Per-seat pricing?</p>
		<span class="chungus__snap" aria-hidden="true">SNAP</span>
		<canvas class="chungus__dust" aria-hidden="true"></canvas>
		<div class="chungus__flash" aria-hidden="true"></div>
		<button class="chungus__close" type="button" aria-label="Close">&times;</button>`;
}

/** Pin an overlay element to a point on screen (the head, the hand). */
function pin(el, point) {
	el.style.left = `${point.x}px`;
	el.style.top = `${point.y}px`;
}

export function initChungus() {
	const trigger = document.getElementById('chungus');
	// An easter egg: the gauntlet only shows up now and then.
	const forced = new URLSearchParams(location.search).has('chungus');
	if (trigger && (forced || Math.random() < APPEARANCE_CHANCE)) trigger.hidden = false;
	/** @type {HTMLElement | null} */
	let root = null;
	let run = 0;
	/** @type {ReturnType<typeof createSound>} */
	let sound;
	/** @type {Promise<any>} */
	let scene = Promise.resolve(null);

	function close() {
		if (!root) return;
		run++;
		sound.close();
		void scene.then((s) => s?.dispose());
		root.remove();
		root = null;
		document.body.classList.remove('chungus-open');
		document.removeEventListener('keydown', onKey);
		trigger?.focus();
	}

	function onKey(event) {
		if (event.key === 'Escape') close();
	}

	async function play(stage) {
		const token = ++run;
		const fast = reducedMotion();
		const caption = /** @type {HTMLElement} */ (stage.querySelector('.chungus__caption'));
		const guard = (promise) =>
			Promise.resolve(promise).then((value) => {
				if (token !== run) throw ABORT;
				return value;
			});
		const wait = (ms) => guard(new Promise((r) => setTimeout(r, fast ? ms / 3 : ms)));
		const say = (text) => (caption.textContent = text);
		// Phases accumulate, so CSS can say "from the snap onwards" with one class.
		const phase = (name) => stage.classList.add(`is-${name}`);
		const stomp = () => {
			stage.classList.remove('is-thud');
			void stage.offsetWidth;
			stage.classList.add('is-thud');
			sound.stomp();
		};
		scene = import('./chungus-scene.js')
			.then((m) =>
				m.createScene(/** @type {HTMLElement} */ (stage.querySelector('.chungus__world')), {
					fast,
					onStep: stomp,
				}),
			)
			.catch(() => null);

		try {
			phase('old');
			say('Somewhere, a licence renews itself.');
			const [s] = await guard(Promise.all([scene, wait(1400)]));
			phase('walk');
			say('Something very large approaches.');
			await guard(s ? s.walkIn() : wait(800));
			await guard(s?.face());
			phase('speak');
			say('');
			if (s)
				pin(/** @type {HTMLElement} */ (stage.querySelector('.chungus__bubble')), s.headPoint());
			await wait(1600);
			phase('raise');
			await guard(s?.raise());
			phase('charge');
			sound.hum();
			await guard(s ? s.charge() : wait(1200));
			if (s) pin(/** @type {HTMLElement} */ (stage.querySelector('.chungus__snap')), s.handPoint());
			phase('snap');
			sound.snap();
			await guard(s?.snap());
			await wait(900);
			phase('dust');
			say('...');
			void s?.release();
			await guard(fast ? wait(600) : disintegrate(stage, OLD_SUITE));
			phase('new');
			void s?.relax();
			say('Perfectly balanced, as all file formats should be.');
		} catch (error) {
			if (error !== ABORT) throw error;
		}
	}

	function open() {
		if (root) return;
		sound = createSound();
		root = document.createElement('div');
		root.className = 'chungus';
		root.setAttribute('role', 'dialog');
		root.setAttribute('aria-modal', 'true');
		root.setAttribute('aria-label', 'Chungus');
		root.innerHTML = markup();
		document.body.append(root);
		document.body.classList.add('chungus-open');
		document.addEventListener('keydown', onKey);
		for (const button of root.querySelectorAll('.chungus__close, [data-close]'))
			button.addEventListener('click', close);
		/** @type {HTMLElement | null} */ (root.querySelector('.chungus__close'))?.focus();
		void play(root);
	}

	trigger?.addEventListener('click', open);
	// Any link inside (an app, or back to the suite) changes the route.
	window.addEventListener('hashchange', close);
}
