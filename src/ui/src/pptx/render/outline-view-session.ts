/**
 * The Outline view's edit session: the one place an outline edit or key press becomes a
 * committed deck and a caret hand-off.
 *
 * The rules (what a row is, what Tab does, which gesture makes a slide) were already shared in
 * `outline-view` and `outline-view-edit`; each binding still kept its own copy of the shell
 * around them (refuse on a read-only deck, apply, skip a no-op, commit, then move the caret to
 * the row the edit names once that row exists). The copies looked for the row three ways: a
 * `CSS.escape`d attribute selector (React, Vue, Angular), an attribute scan (Svelte) and a map
 * of inputs (vanilla). This session is that shell, written once, and {@link focusOutlineRow} is
 * the one lookup.
 *
 * A binding creates one session per open outline and calls {@link OutlineViewSession.restoreFocus}
 * after its next render (an effect, `nextTick`, a microtask or straight after a synchronous
 * repaint), because the row that should take the caret may only appear once the new deck is
 * drawn.
 *
 * @module render/outline-view-session
 */
import type { PptxSlide } from 'ooxml-core/pptx';

import { OUTLINE_ROW_ATTR } from './outline-view';
import { applyOutlineEdit, mapOutlineKey } from './outline-view-edit';
import type { OutlineEdit, OutlineKeyInput } from './outline-view-edit';

export interface OutlineViewSessionDeps {
	/** The editable (template-free) deck as it is now. */
	slides(): readonly PptxSlide[];
	/** The canvas a title element the outline creates is laid out against. */
	canvasSize(): { width: number; height: number };
	/** False for a read-only deck, which makes every edit a no-op. */
	canEdit(): boolean;
	/**
	 * Replace the deck and make `activeSlideIndex` current, through the binding's own undoable
	 * whole-deck commit, so an outline edit is one ordinary undo step.
	 */
	commit(slides: PptxSlide[], activeSlideIndex: number): void;
}

/** The part of a keyboard event the session reads. */
export interface OutlineViewKeyEvent extends OutlineKeyInput {
	preventDefault(): void;
}

export interface OutlineViewSession {
	/** Apply one edit; returns whether the deck changed (and was committed). */
	run(edit: OutlineEdit): boolean;
	/** The row text field's value changed. */
	setText(rowKey: string, text: string): boolean;
	/**
	 * A key press inside row `rowKey`. Tab is prevented (it would leave the outline) and Enter is
	 * prevented (it would submit a surrounding form), as the shared keymap says.
	 */
	handleKey(event: OutlineViewKeyEvent, rowKey: string): void;
	/** The row the last edit asked the caret to land on, until it is restored. */
	pendingFocusKey(): string | null;
	/**
	 * Move the caret to the pending row inside `container`, once; a no-op when nothing is pending
	 * or the row is not drawn yet (the request is then dropped, as a stale one would be).
	 */
	restoreFocus(container: ParentNode | null | undefined): void;
}

/**
 * Focus the text field of row `key` inside `container`, caret at the end of the line. Rows are
 * matched by comparing the attribute, not with a selector: a row key carries `|` and spaces,
 * and escaping those into a selector is a hazard for no gain over a scan of the outline.
 */
export function focusOutlineRow(container: ParentNode, key: string): boolean {
	const target = [...container.querySelectorAll<HTMLInputElement>(`[${OUTLINE_ROW_ATTR}]`)].find(
		(input) => input.getAttribute(OUTLINE_ROW_ATTR) === key,
	);
	if (!target) {
		return false;
	}
	if (target.ownerDocument.activeElement !== target) {
		target.focus();
		target.setSelectionRange(target.value.length, target.value.length);
	}
	return true;
}

export function createOutlineViewSession(deps: OutlineViewSessionDeps): OutlineViewSession {
	let pending: string | null = null;
	const run = (edit: OutlineEdit): boolean => {
		if (!deps.canEdit()) {
			return false;
		}
		const result = applyOutlineEdit(deps.slides(), edit, { canvas: deps.canvasSize() });
		if (!result.changed) {
			return false;
		}
		pending = result.focusKey;
		deps.commit(result.slides, result.activeSlideIndex);
		return true;
	};
	return {
		run,
		setText: (key, text) => run({ type: 'setText', key, text }),
		handleKey(event, rowKey) {
			const { edit, preventDefault } = mapOutlineKey(event, rowKey);
			if (preventDefault) {
				event.preventDefault();
			}
			if (edit) {
				run(edit);
			}
		},
		pendingFocusKey: () => pending,
		restoreFocus(container) {
			const key = pending;
			if (key === null || !container) {
				return;
			}
			pending = null;
			focusOutlineRow(container, key);
		},
	};
}
