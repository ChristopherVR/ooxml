/**
 * The "letter" edition of a release post (edition-letter.css, scripts/release-edition.mjs): each
 * recording gets a control that pauses and plays it, and the page moves as it is read. Headings
 * come up word by word, blocks rise, recordings grow into their frames and a chapter's hairline
 * draws across, each once, as it reaches the window.
 *
 * The motion is only ever visual. Nothing is hidden until everything here has run (the `motion`
 * class goes on last), hidden parts stay in the accessibility tree, a part that is already
 * above the window or that takes focus is shown at once, and a link into the page shows the
 * part it lands on. With reduced motion none of it is turned on. It keeps nothing on the device
 * and makes no request.
 */
const body = document.body;
const reduced = matchMedia('(prefers-reduced-motion: reduce)');

// Recordings. showcase.js plays a GIF while it is in view and shows its poster otherwise, by
// reading `data-animation`; pausing points that at the poster, so the two never disagree.
for (const image of document.querySelectorAll('.post-figure img[data-animation]')) {
	const recording = image.dataset.animation;
	const poster = image.getAttribute('src');
	const frame = document.createElement('div');
	frame.className = 'frame';
	image.before(frame);
	frame.append(image);
	const button = document.createElement('button');
	button.type = 'button';
	const full = document.createElement('a');
	full.href = recording;
	full.textContent = 'Full size';
	const bar = document.createElement('div');
	bar.className = 'recording-bar';
	bar.append(button, full);
	frame.after(bar);
	const show = () => {
		const playing = image.getAttribute('src') === recording;
		// A recording that waits for its place on the screen is still "playing" to the reader.
		const paused = !playing && (reduced.matches || image.dataset.animation === poster);
		button.dataset.state = paused ? 'paused' : 'playing';
		button.textContent = paused ? 'Play recording' : 'Pause recording';
	};
	button.addEventListener('click', () => {
		const pause = button.dataset.state === 'playing';
		image.dataset.animation = pause ? poster : recording;
		image.src = pause ? poster : recording;
		show();
	});
	new MutationObserver(show).observe(image, { attributes: true, attributeFilter: ['src'] });
	reduced.addEventListener('change', show);
	show();
}

/** Wraps each word of a plain heading so it can come up from behind its baseline. */
function split(heading) {
	const text = heading.textContent.trim();
	if (!text) return;
	// Assistive technology reads the heading as one phrase, not as the pieces drawn here.
	heading.setAttribute('aria-label', text);
	const pieces = document.createElement('span');
	pieces.setAttribute('aria-hidden', 'true');
	let index = 0;
	for (const part of text.split(/(\s+)/u)) {
		if (!part) continue;
		if (/^\s+$/u.test(part)) {
			pieces.append(' ');
			continue;
		}
		const mask = document.createElement('span');
		mask.className = 'w';
		const inner = document.createElement('span');
		inner.textContent = part;
		inner.style.setProperty('--i', String(index++));
		mask.append(inner);
		pieces.append(mask);
	}
	heading.replaceChildren(pieces);
}

function startMotion() {
	const mark = (selector, kind) => {
		for (const element of document.querySelectorAll(selector))
			if (!element.closest('[data-rv]')) element.dataset.rv = kind;
	};
	// The opening (title, date, sentence, contents, lede) animates from the stylesheet alone.
	for (const heading of document.querySelectorAll('.chapter-head h2, .feature > h3')) {
		if (heading.children.length) continue;
		split(heading);
		heading.dataset.rv = 'words';
	}
	mark('.chapter-links', 'rise');
	mark('.chapter-head', 'head');
	mark('.post-article .post-figure', 'frame');
	// Lists come in an item at a time; everything else as the block it is.
	mark('.post-article :is(.post-section, .feature-text) > :is(ul, ol) > li', 'rise');
	mark(
		'.post-article :is(.post-section, .feature-text, .feature-aside) > :is(p, figure, .table-wrap, .migration-pair, blockquote, .callout, h3, h4), details.changes, .changes-toggle',
		'rise',
	);
	const parts = [...document.querySelectorAll('[data-rv]')];
	const show = (element) => element.classList.add('in');
	// What is already on the screen (a link into the page, a reload half way down) was painted
	// before this ran: it stays as it is rather than hiding to make an entrance.
	for (const element of parts) if (element.getBoundingClientRect().top < innerHeight) show(element);
	const observer = new IntersectionObserver(
		(entries) => {
			let order = 0;
			for (const entry of entries) {
				const element = entry.target;
				// Already read past (a link or a reload landed below it): show it without ceremony.
				const above = entry.boundingClientRect.bottom <= 0;
				if (!entry.isIntersecting && !above) continue;
				if (entry.isIntersecting && !above)
					element.style.setProperty('--d', `${Math.min(order++, 6) * 80}ms`);
				show(element);
				observer.unobserve(element);
			}
		},
		{ rootMargin: '0px 0px -10% 0px' },
	);
	for (const element of parts) if (!element.classList.contains('in')) observer.observe(element);
	// The part a link lands on, and anything that takes focus, is never left waiting.
	const reveal = (target) => {
		const scope = target?.closest?.('.feature, .chapter-head, .post-lede, details.changes');
		for (const element of [target?.closest?.('[data-rv]'), scope])
			if (element?.dataset?.rv) show(element);
		for (const element of scope?.querySelectorAll('[data-rv]') ?? []) show(element);
	};
	const landed = () => reveal(document.getElementById(decodeURIComponent(location.hash.slice(1))));
	addEventListener('hashchange', landed);
	document.addEventListener('focusin', (event) => reveal(event.target));
	if (location.hash) landed();
	body.classList.add('motion');
}

const stopMotion = () => body.classList.remove('motion');
if ('IntersectionObserver' in window && !reduced.matches) {
	try {
		startMotion();
	} catch {
		stopMotion();
	}
}
// A reader who turns motion off while the page is open gets the still page at once.
reduced.addEventListener('change', () => {
	if (reduced.matches) stopMotion();
});
