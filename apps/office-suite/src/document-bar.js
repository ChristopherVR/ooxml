/**
 * Simple view: the suite's own short command bar over a document whose ribbon is held back. It
 * shows the file, undo and redo, bold, italic and underline (Word, Excel), the slide or page in
 * view (PowerPoint, Visio), and one button that brings each app's ribbon back. Everything here is
 * suite chrome; the editors are driven through the `view` object `mountEditor` returns and keep
 * their own UI untouched when the ribbon is on.
 */
import { $, badge, escape, task } from './ui.js';
import { icon } from './icons.js';
import { onSettings, settings, updateSettings } from './settings.js';

/** `Quarterly review.pptx` as its stem and its extension, so the bar can mute the extension. */
export function splitFileName(name) {
	const dot = name.lastIndexOf('.');
	return dot > 0 ? [name.slice(0, dot), name.slice(dot)] : [name, ''];
}

/** What the position control shows for a `view.state().position`. */
export function positionLabel({ unit, index, count }) {
	return count > 0 ? `${unit} ${Math.min(index, count - 1) + 1} / ${count}` : '';
}

/**
 * `aria-pressed` for a formatting button: the editors report `null` when the selection mixes
 * bold and plain text, which a toggle button states as "mixed".
 */
export function pressedState(value) {
	return value === null ? 'mixed' : String(Boolean(value));
}

/**
 * `active()` is the open document's id while an editor is in view, else null. Returns the
 * function that redraws the bar.
 */
export function mountDocumentBar({ sessions, active }) {
	const view = () => sessions.sessions.get(active())?.editor?.view;
	$('doc-undo').innerHTML = icon('undo');
	$('doc-redo').innerHTML = icon('redo');
	$('doc-previous').innerHTML = icon('back');
	$('doc-next').innerHTML = icon('forward');
	// A press on the bar must not take the keyboard out of the document: the caret and the
	// selection the command applies to stay where they are.
	$('doc-view').addEventListener('mousedown', (event) => {
		if (event.target.closest('button')) event.preventDefault();
	});
	$('doc-undo').onclick = () => task(() => view()?.undo());
	$('doc-redo').onclick = () => task(() => view()?.redo());
	$('doc-format').addEventListener('click', (event) => {
		const key = event.target.closest('[data-format]')?.dataset.format;
		if (key) task(() => view()?.toggleFormat?.(key)).then(render);
	});
	const step = (delta) =>
		task(() => {
			const position = view()?.state().position;
			if (!position) return;
			const next = position.index + delta;
			if (next >= 0 && next < position.count) view().go(next);
		});
	$('doc-previous').onclick = () => step(-1);
	$('doc-next').onclick = () => step(1);
	$('layout-toggle').onclick = () =>
		updateSettings({ layout: settings().layout === 'simple' ? 'office' : 'simple' });
	const focus = $('focus-toggle');
	focus.onclick = () => {
		document.body.dataset.focus = String(document.body.dataset.focus !== 'true');
		render();
	};

	let shownTitle = '';
	function render() {
		const simple = settings().layout === 'simple';
		document.body.dataset.layout = simple ? 'simple' : 'office';
		$('layout-toggle').innerHTML = simple
			? `${icon('ribbon')}<span>Show ribbon</span>`
			: `${icon('page')}<span>Simple view</span>`;
		$('layout-toggle').title = simple
			? "Show each app's ribbon"
			: 'Hide the ribbon and keep a short command bar';
		const focused = document.body.dataset.focus === 'true';
		focus.innerHTML = icon(focused ? 'collapse' : 'expand');
		focus.setAttribute('aria-pressed', String(focused));
		focus.setAttribute('aria-label', 'Focus on the document');
		focus.title = focused ? 'Show the workspace again' : 'Focus on the document';
		const session = sessions.sessions.get(active());
		if (!simple || !session) return;
		const [stem, extension] = splitFileName(session.doc.name);
		const title = `${badge(session.doc.kind)}<span class="doc-name">${escape(stem)}<span class="doc-extension">${escape(extension)}</span></span>`;
		// The bar redraws on every selection change; the title only when it differs.
		if (shownTitle !== title) {
			shownTitle = title;
			$('doc-title').innerHTML = title;
			$('doc-title').title = session.doc.name;
		}
		const state = session.editor?.view?.state();
		$('doc-view').hidden = !state;
		if (!state) return;
		$('doc-undo').disabled = !state.canUndo;
		$('doc-redo').disabled = !state.canRedo;
		$('doc-format').hidden = !state.format;
		for (const button of $('doc-format').querySelectorAll('[data-format]')) {
			button.disabled = !state.format?.enabled;
			button.setAttribute('aria-pressed', pressedState(state.format?.[button.dataset.format]));
		}
		const position = state.position;
		$('doc-nav').hidden = !position || position.count < 1;
		if (!position) return;
		const unit = position.unit.toLowerCase();
		$('doc-position').textContent = positionLabel(position);
		for (const [id, label, off] of [
			['doc-previous', `Previous ${unit}`, position.index <= 0],
			['doc-next', `Next ${unit}`, position.index >= position.count - 1],
		]) {
			$(id).disabled = off;
			$(id).title = label;
			$(id).setAttribute('aria-label', label);
		}
	}

	// The editors follow the setting themselves (editors.js); the bar only redraws.
	onSettings((_next, keys) => {
		if (keys.includes('layout')) render();
	});
	// Selection changes arrive in bursts (every key press in Word): one redraw per frame.
	let queued = false;
	sessions.moved = () => {
		if (queued) return;
		queued = true;
		requestAnimationFrame(() => {
			queued = false;
			render();
		});
	};
	render();
	return render;
}
