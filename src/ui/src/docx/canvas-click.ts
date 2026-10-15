import { TextSelection } from 'prosemirror-state';
import type { EditorCore } from './editor-core';

const clamp = (value: number, low: number, high: number) =>
	Math.min(Math.max(value, low), Math.max(low, high));

/**
 * A press on the bare canvas (the gutter beside the pages, or around the Draft paper) lands on no
 * text, so the browser would focus the canvas and leave typing nowhere. Word puts the caret at the
 * nearest text instead; do the same in both surfaces.
 */
export function attachCanvasClick(core: EditorCore): void {
	const canvas = core.shell.canvas;
	if (!canvas) return;
	const editable = () => {
		const view = core.view;
		return view &&
			!view.isDestroyed &&
			view.editable &&
			!core.readOnly &&
			core.targetView() === view
			? view
			: undefined;
	};
	canvas.addEventListener('pointerdown', (event) => {
		const view = editable();
		if (!view || event.button !== 0) return;
		if (event.target !== canvas && event.target !== core.shell.paper) return;
		const frame = canvas.getBoundingClientRect();
		// A press on the scrollbar also targets the canvas; leave it to the browser.
		if (
			event.clientX - frame.left >= canvas.clientWidth ||
			event.clientY - frame.top >= canvas.clientHeight
		)
			return;
		if (core.pages.viewMode === 'print') {
			if (core.shell.printEditing?.selectNearest(event.clientX, event.clientY))
				event.preventDefault();
			return;
		}
		const box = view.dom.getBoundingClientRect();
		const found = view.posAtCoords({
			left: clamp(event.clientX, box.left + 1, box.right - 1),
			top: clamp(event.clientY, box.top + 1, box.bottom - 1),
		});
		if (!found) return;
		event.preventDefault();
		view.dispatch(
			view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(found.pos))),
		);
		view.focus();
	});
	// The canvas is focusable only so a read-only document has somewhere to take focus. A touch tap
	// focuses the tapped element when the gesture ends, after pointerdown; when the press placed the
	// caret, hand that focus back to the editor. Read-only content keeps its deliberate blur.
	let armed = 0;
	canvas.addEventListener(
		'pointerdown',
		(event) => {
			const target = event.target instanceof Element ? event.target : null;
			armed =
				editable() && event.button === 0 && !target?.closest('[data-read-only="true"]')
					? Date.now() + 1000
					: 0;
		},
		{ capture: true },
	);
	canvas.addEventListener('focus', (event) => {
		if (event.target === canvas && Date.now() < armed) editable()?.focus();
	});
}
