import { closeHistory } from 'prosemirror-history';
import { wordYjsPluginKey } from 'ooxml-core/docx/ui';
import type { EditorView } from 'prosemirror-view';
import { selectionPaneFrame } from '../chrome/selection-pane';
import { drawingPositions, isArrangeDrawing, selectDrawing } from './arrange-selection';
import type { EditorCore } from './editor-core';
import { translate, translateTemplate } from './localization';
import { parsePlacement } from './print-float-locate';
import { focusView } from './focus-view';

export function storyDrawings(view: EditorView) {
	const drawings: { pos: number; title: string; height: number; behind: boolean }[] = [];
	view.state.doc.descendants((node, pos) => {
		if (!isArrangeDrawing(node)) return true;
		const placement = parsePlacement(node.attrs.placement);
		drawings.push({
			pos,
			title: String(node.attrs.drawingName || node.attrs.title || node.attrs.altText || ''),
			height: placement?.relativeHeight ?? 251658240 + drawings.length,
			behind: Boolean(placement?.behindText),
		});
		return false;
	});
	return drawings;
}

const panes = new WeakMap<EditorCore, (view: EditorView) => void>();

export function toggleArrangePane(core: EditorCore, view: EditorView): void {
	const pane = panes.get(core);
	if (pane) return pane(view);
	const dock = core.shell.canvas?.parentElement;
	if (!dock) return;
	let target = view;
	const frame = selectionPaneFrame(dock.ownerDocument, () => {
		if (!target.isDestroyed) focusView(target);
	});
	const { element, heading, close, list, hide } = frame;
	element.className = 'dve-selection-pane';
	// Keep native input and button behavior while isolating the pane from editor shortcuts.
	for (const type of ['keydown', 'keypress', 'keyup'])
		element.addEventListener(type, (event) => event.stopPropagation());
	dock.append(element);
	let lastDoc: EditorView['state']['doc'] | undefined;
	let drawings: ReturnType<typeof storyDrawings> = [];
	let signature = '';
	let pending: ReturnType<typeof setTimeout> | undefined;
	function refresh() {
		if (element.hidden) return;
		if (!element.isConnected) return;
		const live = core.targetView();
		if (!live || live.isDestroyed) return hide();
		target = live;
		if (live.state.doc !== lastDoc) {
			lastDoc = live.state.doc;
			drawings = storyDrawings(live);
		}
		const selected = drawingPositions(live.state);
		const next = JSON.stringify([core.locale, drawings]);
		if (signature === next) {
			for (const button of list.querySelectorAll<HTMLButtonElement>(
				'button[data-drawing-position]',
			))
				button.setAttribute(
					'aria-pressed',
					String(selected.includes(Number(button.dataset.drawingPosition))),
				);
			return;
		}
		signature = next;
		const title = translate(core.locale, 'Selection Pane');
		heading.textContent = title;
		element.setAttribute('aria-label', title);
		close.setAttribute('aria-label', translate(core.locale, 'Close'));
		const focused = core.element.shadowRoot?.activeElement;
		const active = focused instanceof HTMLElement ? focused.dataset.drawingPosition : undefined;
		list.replaceChildren(
			...drawings
				.slice(0, 1000)
				.sort((a, b) => Number(a.behind) - Number(b.behind) || b.height - a.height)
				.map(({ pos, title }, index) => {
					const row = element.ownerDocument.createElement('li');
					const button = element.ownerDocument.createElement('button');
					button.type = 'button';
					button.dataset.drawingPosition = String(pos);
					button.textContent =
						title || translateTemplate(core.locale, 'Drawing {number}', { number: index + 1 });
					button.setAttribute('aria-pressed', String(selected.includes(pos)));
					button.addEventListener('click', (event) => {
						if (live !== core.targetView()) return;
						// Flush any pending pane refresh before acting on the displayed row.
						refresh();
						if (!button.isConnected) return;
						const multiple = event.ctrlKey || event.metaKey || event.shiftKey;
						selectDrawing(live, pos, multiple);
						refresh();
					});
					const rename = () => {
						if (live !== core.targetView() || !live.editable || core.readOnly) return;
						const node = live.state.doc.nodeAt(pos);
						if (!isArrangeDrawing(node) || !button.isConnected) return;
						const input = element.ownerDocument.createElement('input');
						input.type = 'text';
						input.dataset.drawingPosition = String(pos);
						input.setAttribute('aria-label', translate(core.locale, 'Drawing name'));
						input.value = button.textContent ?? '';
						input.maxLength = 255;
						let editing = true;
						const finish = (save: boolean, returnFocus: boolean) => {
							if (!editing || !input.isConnected) return;
							editing = false;
							button.hidden = false;
							if (returnFocus) button.focus({ preventScroll: true });
							input.remove();
							const name = input.value.trim();
							if (
								save &&
								name &&
								name !== button.textContent &&
								live === core.targetView() &&
								live.editable &&
								!core.readOnly &&
								live.state.doc.nodeAt(pos) === node
							) {
								const history = wordYjsPluginKey.getState(live.state);
								history?.stopCapturing();
								live.dispatch(
									closeHistory(live.state.tr.setNodeAttribute(pos, 'drawingName', name)),
								);
								history?.stopCapturing();
								refresh();
							}
						};
						input.addEventListener('blur', () => finish(true, false));
						input.addEventListener('keydown', (event) => {
							if (event.isComposing) {
								event.stopPropagation();
								return;
							}
							if (event.key !== 'Enter' && event.key !== 'Escape') return;
							event.preventDefault();
							event.stopPropagation();
							finish(event.key === 'Enter', true);
						});
						button.after(input);
						button.hidden = true;
						input.focus();
						input.select();
					};
					button.addEventListener('dblclick', rename);
					button.addEventListener('keydown', (event) => {
						if (event.key !== 'F2') return;
						event.preventDefault();
						rename();
					});
					row.append(button);
					return row;
				}),
		);
		if (active !== undefined)
			list.querySelector<HTMLElement>(`[data-drawing-position="${active}"]`)?.focus();
	}
	core.element.addEventListener('editor-state-change', () => {
		if (element.hidden || !element.isConnected) return;
		clearTimeout(pending);
		pending = setTimeout(refresh, 120);
	});
	const toggle = (next: EditorView) => {
		if (!element.hidden) return hide();
		target = next;
		element.hidden = false;
		refresh();
		close.focus();
	};
	panes.set(core, toggle);
	toggle(view);
}
