const preference = matchMedia('(prefers-reduced-motion: reduce)');
const dialog = document.createElement('dialog');
dialog.className = 'preview-dialog';
dialog.setAttribute('aria-label', 'Enlarged product preview');
dialog.innerHTML = '<button type="button" class="preview-close">Close</button><img alt="">';
document.body.append(dialog);
const enlarged = dialog.querySelector('img');
const closePreview = dialog.querySelector('.preview-close');
let opener;

document.addEventListener('click', (event) => {
	const button = event.target.closest?.('[data-preview]');
	if (!button || dialog.open) return;
	const image = button.querySelector('img');
	opener = button;
	enlarged.src = image.dataset.animation || image.currentSrc || image.src;
	enlarged.alt = image.alt;
	dialog.showModal();
	dialog.dispatchEvent(new Event('preview-open'));
});
closePreview.addEventListener('click', () => dialog.close());
dialog.addEventListener('click', (event) => {
	if (event.target === dialog) dialog.close();
});
dialog.addEventListener('close', () => {
	enlarged.removeAttribute('src');
	opener?.focus({ preventScroll: true });
});

function watchView(element, setVisible) {
	if (!('IntersectionObserver' in window)) {
		setVisible(true);
		return;
	}
	new IntersectionObserver(([entry]) => setVisible(entry.isIntersecting), {
		rootMargin: '-8% 0px -16% 0px',
	}).observe(element);
}

function setAnimation(element, enabled) {
	for (const image of element.querySelectorAll('img[data-animation]')) {
		const source = enabled ? image.dataset.animation : image.dataset.poster;
		if (image.getAttribute('src') !== source) image.src = source;
	}
}

for (const showcase of document.querySelectorAll('[data-showcase]')) {
	const stage = showcase.querySelector('.showcase-stage');
	const slides = [...showcase.querySelectorAll('.showcase-panel')];
	const buttons = [...showcase.querySelectorAll('[data-slide]')];
	const controls = showcase.querySelector('.showcase-controls');
	const notes = [...showcase.querySelectorAll('[data-note]')];
	const shell = showcase.dataset.showcase === 'shell';
	const delay = shell ? 960 : showcase.dataset.showcase === 'performance' ? 9200 : 6500;
	let visible = false;
	let hovering = false;
	let pinned = false;
	let started = false;
	let current = 0;
	let requested = 0;
	let timer;
	let fade;
	let outgoing;
	let shellMoves = 0;

	function finishFade() {
		clearTimeout(fade);
		if (!outgoing) return;
		outgoing.hidden = true;
		outgoing.inert = false;
		delete outgoing.dataset.exiting;
		delete slides[current].dataset.entering;
		outgoing = undefined;
	}

	controls.hidden = false;

	function animate(panel) {
		delete panel.dataset.active;
		// Restart the example when someone picks the same tab again.
		void panel.offsetWidth;
		panel.dataset.active = '';
	}

	function canRotate() {
		return (
			visible &&
			!pinned &&
			!hovering &&
			!document.hidden &&
			!preference.matches &&
			!dialog.open &&
			!showcase.contains(document.activeElement) &&
			slides.length > 1 &&
			(!shell || shellMoves < slides.length - 1)
		);
	}

	async function select(index, manual = false) {
		requested = index;
		const panel = slides[index];
		const image = panel.querySelector('img');
		if (image && !image.dataset.animation) {
			image.loading = 'eager';
			await image.decode().catch(() => {});
		}
		if (requested !== index || (!manual && !canRotate())) return;
		if (index !== current) {
			finishFade();
			const previous = slides[current];
			setAnimation(previous, false);
			if (shell && !preference.matches) {
				outgoing = previous;
				previous.inert = true;
				previous.dataset.exiting = '';
				panel.dataset.entering = '';
				fade = setTimeout(finishFade, 430);
			} else previous.hidden = true;
			if (notes[current]) notes[current].hidden = true;
			buttons[current].setAttribute('aria-pressed', 'false');
			current = index;
			panel.hidden = false;
			if (notes[current]) notes[current].hidden = false;
			buttons[current].setAttribute('aria-pressed', 'true');
			if (shell && !manual) shellMoves++;
		}
		animate(panel);
		schedule();
	}

	function schedule() {
		clearTimeout(timer);
		const motion = visible && !document.hidden && !preference.matches && !dialog.open;
		setAnimation(slides[current], motion);
		if (canRotate()) timer = setTimeout(() => void select((current + 1) % slides.length), delay);
	}

	for (const [index, button] of buttons.entries())
		button.addEventListener('click', () => {
			pinned = true;
			clearTimeout(timer);
			void select(index, true);
		});
	showcase.addEventListener('focusin', schedule);
	showcase.addEventListener('focusout', () => queueMicrotask(schedule));
	showcase.addEventListener('pointerenter', (event) => {
		if (event.pointerType === 'mouse') hovering = true;
		schedule();
	});
	showcase.addEventListener('pointerleave', () => {
		hovering = false;
		schedule();
	});
	dialog.addEventListener('preview-open', schedule);
	dialog.addEventListener('close', schedule);
	document.addEventListener('visibilitychange', schedule);
	preference.addEventListener('change', schedule);
	watchView(stage, (inView) => {
		visible = inView;
		if (visible && !started) {
			started = true;
			animate(slides[current]);
			if (shell) for (const panel of slides) panel.querySelector('img').loading = 'eager';
		}
		schedule();
	});
}

for (const figure of document.querySelectorAll('figure img[data-animation]')) {
	const poster = figure.getAttribute('src');
	watchView(figure, (visible) => {
		const src = visible && !preference.matches ? figure.dataset.animation : poster;
		if (figure.getAttribute('src') !== src) figure.src = src;
	});
}

for (const shortcut of document.querySelectorAll('.flow-shortcut')) {
	const panel = shortcut.closest('.flow-copy');
	const steps = [...panel.querySelectorAll('.flow-steps li')];
	let index = 0;
	const move = (direction) => {
		steps[index].removeAttribute('aria-current');
		index = (index + direction + steps.length) % steps.length;
		steps[index].setAttribute('aria-current', 'step');
	};
	for (const key of shortcut.querySelectorAll('[data-step-direction]'))
		key.addEventListener('click', () => move(Number(key.dataset.stepDirection)));
	shortcut.addEventListener('keydown', (event) => {
		if (event.key !== 'F6') return;
		event.preventDefault();
		move(event.shiftKey ? -1 : 1);
	});
}

for (const trigger of document.querySelectorAll('[data-unlock]')) {
	const flow = trigger.closest('[data-gated]');
	trigger.addEventListener('click', () => {
		const unlocked = flow.toggleAttribute('data-unlocked');
		trigger.textContent = unlocked ? 'Block content' : 'Enable content';
	});
}

for (const scene of document.querySelectorAll('.security-scene')) {
	const choices = [...scene.querySelectorAll('[data-security-choice]')];
	const cases = [...scene.querySelectorAll('[data-security-case]')];
	for (const [i, button] of choices.entries())
		button.addEventListener('click', () => {
			for (const [index, choice] of choices.entries()) {
				choice.setAttribute('aria-pressed', String(index === i));
				cases[index].hidden = index !== i;
			}
		});
}
