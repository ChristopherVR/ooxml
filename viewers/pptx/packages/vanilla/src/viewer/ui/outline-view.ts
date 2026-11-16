import type { CanvasSize, OutlineRow } from 'ooxml-ui/pptx';
import {
	buildOutline,
	createOutlineViewSession,
	listenForOutlineViewDismiss,
	OUTLINE_LEVEL_ATTR,
	OUTLINE_ROW_ATTR,
	OUTLINE_SLIDE_ATTR,
	OUTLINE_VIEW_ATTR,
} from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

import type { Translator } from '../i18n';
import { createEl } from '../render';
import { makeButton } from './controls';

/** Indent per outline level, in pixels. Level 0 (a title) sits flush left. */
const INDENT_PX = 22;

export interface OutlineViewOptions {
	slides: readonly PptxSlide[];
	/** Laid out against by a title element the outline has to create. */
	canvasSize: CanvasSize;
	/** False renders every row `readonly` and makes every edit a no-op. */
	canEdit: boolean;
	/**
	 * Follow the deck while the pane is open.
	 *
	 * Undo/redo and a collaborator's edit both replace the deck underneath an
	 * open outline, and a pane still showing its opening snapshot would write
	 * that stale deck back on the next keystroke.
	 */
	subscribe?(listener: (slides: readonly PptxSlide[]) => void): () => void;
	/**
	 * Hand the edited deck to the viewer's own whole-deck commit
	 * (`EditorController.commitSlides`), which is what puts an outline edit in
	 * the same history as a slide reorder or delete.
	 */
	commit(slides: PptxSlide[], activeSlideIndex: number): void;
	/** The user asked to go back to Normal view. */
	onClose(): void;
}

export interface OutlineViewHandle {
	el: HTMLElement;
	/** Tear down: drop the store subscription and the node. */
	close(): void;
}

/**
 * The row set as the DOM currently holds it.
 *
 * Keys AND levels, because a promote/demote leaves every key in place but has
 * to repaint the indent.
 */
function rowSignature(rows: readonly OutlineRow[]): string {
	return rows.map((row) => `${row.key}@${row.level}`).join('\u0000');
}

/**
 * PowerPoint's Outline view: the deck as an editable indented text document.
 *
 * One row per slide title at the left margin, that slide's body lines stepped
 * in beneath it. Every rule (what a row is, what Tab does, which gesture makes
 * a slide) and the edit session (`createOutlineViewSession`) come from
 * `ooxml-ui/pptx`, so the five bindings cannot drift; this module is the
 * markup and the listeners, and nothing else.
 *
 * Each row is a real `<input>` rather than a contenteditable: caret placement,
 * IME commit and native undo are exactly what a one-line text field already
 * gets right, and the viewer's global shortcut handler (`ui/keyboard.ts`)
 * already declines any key that started inside an INPUT, so typing here cannot
 * reach the editor underneath.
 */
export function openOutlineViewOverlay(
	doc: Document,
	host: HTMLElement,
	t: Translator,
	options: OutlineViewOptions,
): OutlineViewHandle {
	host.querySelector(`[${OUTLINE_VIEW_ATTR}]`)?.remove();

	let slides = options.slides;
	let signature = '';
	let closed = false;
	// The edited row's input, kept across a keystroke so the caret stays put.
	const inputs = new Map<string, HTMLInputElement>();

	const el = createEl(doc, 'section', 'pptxv-outline-view');
	el.setAttribute(OUTLINE_VIEW_ATTR, 'true');
	el.setAttribute('role', 'region');
	el.setAttribute('aria-label', t('pptx.view.outlineView'));

	const header = createEl(doc, 'div', 'pptxv-outline-view-header');
	const heading = createEl(doc, 'span', 'pptxv-outline-view-title');
	heading.textContent = t('pptx.view.outlineView');
	const hint = createEl(doc, 'span', 'pptxv-outline-view-hint');
	hint.textContent = t('pptx.outline.hint');
	const exit = makeButton(doc, {
		label: t('pptx.statusBar.normalView'),
		icon: 'close',
		className: 'pptxv-outline-view-btn',
		onClick: () => {
			close();
			options.onClose();
		},
	});
	header.append(heading, hint, exit.btn);
	const list = createEl(doc, 'div', 'pptxv-outline-view-list');
	el.append(header, list);

	const session = createOutlineViewSession({
		slides: () => slides,
		canvasSize: () => options.canvasSize,
		canEdit: () => options.canEdit,
		commit: (next, activeSlideIndex) => {
			slides = next;
			options.commit(next, activeSlideIndex);
			render();
			// The repaint is synchronous, so the row the edit names is already drawn.
			session.restoreFocus(list);
		},
	});

	const buildRow = (row: OutlineRow): HTMLElement => {
		const line = createEl(doc, 'div', 'pptxv-outline-view-row', {
			paddingLeft: `${row.level * INDENT_PX}px`,
		});
		const number = createEl(doc, 'span', 'pptxv-outline-view-number');
		// Drawn only on a slide's first row, which is always its title row, so the
		// pane reads as a list of slides rather than one wall of lines.
		number.textContent = row.kind === 'title' ? String(row.slideIndex + 1) : '';
		const input = doc.createElement('input');
		input.type = 'text';
		input.className = `pptxv-outline-view-input${row.kind === 'title' ? ' is-title' : ''}`;
		input.value = row.text;
		input.readOnly = !options.canEdit;
		input.setAttribute(OUTLINE_ROW_ATTR, row.key);
		input.setAttribute(OUTLINE_SLIDE_ATTR, String(row.slideIndex + 1));
		input.setAttribute(OUTLINE_LEVEL_ATTR, String(row.level));
		input.setAttribute(
			'aria-label',
			t(row.kind === 'title' ? 'pptx.outline.titleLine' : 'pptx.outline.bodyLine'),
		);
		input.addEventListener('input', () => session.setText(row.key, input.value));
		input.addEventListener('keydown', (event) => session.handleKey(event, row.key));
		inputs.set(row.key, input);
		line.append(number, input);
		return line;
	};

	/**
	 * Repaint, replacing nodes only when the row set actually moved.
	 *
	 * A keystroke leaves the same rows in place and the edited input already
	 * holds the text, so rebuilding would destroy the very field being typed into
	 * and drop the caret to the end of the line on every character.
	 */
	function render(): void {
		const rows = buildOutline(slides);
		const next = rowSignature(rows);
		if (next === signature) {
			for (const row of rows) {
				const input = inputs.get(row.key);
				if (input && input.value !== row.text) {
					input.value = row.text;
				}
			}
			return;
		}
		signature = next;
		inputs.clear();
		list.replaceChildren(...rows.map((row) => buildRow(row)));
	}

	const stop = options.subscribe?.((next) => {
		if (!closed && next !== slides) {
			slides = next;
			render();
		}
	});

	function close(): void {
		if (closed) {
			return;
		}
		closed = true;
		stop?.();
		stopDismiss();
		el.remove();
	}

	render();
	host.appendChild(el);
	// Escape returns to Normal view, like the header's close button, wherever
	// focus is (it stays on the ribbon button that opened the view).
	const stopDismiss = listenForOutlineViewDismiss(doc, () => {
		close();
		options.onClose();
	});

	return { el, close };
}
