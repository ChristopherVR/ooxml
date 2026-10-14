import type { SlideBackgroundActions } from '../../editor/editor-background-actions';
import type { EditActions } from '../../editor/editor-edit-ops';
import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { makeButton, makeColorControl } from '../controls';
import { createInspectorCheckbox } from '../inspector/controls-extra';

export interface FormatBackgroundPanel {
	el: HTMLElement;
	toggle(): void;
	setOpen(open: boolean): void;
	isOpen(): boolean;
	setEditable(editable: boolean): void;
}

/**
 * A simple docked "Format Background" panel (solid colour fill only, matching
 * the docked-panel idiom of `find-replace-panel.ts`
 * rather than React's inspector-pane version). Applies immediately on colour
 * pick via {@link SlideBackgroundActions}; the panel doesn't reflect the
 * current slide's existing background (it's a write-only applier, same as
 * the Transitions/Animations galleries).
 */
export function createFormatBackgroundPanel(
	doc: Document,
	t: Translator,
	actions: Pick<
		SlideBackgroundActions,
		'setSlideBackgroundColor' | 'clearSlideBackground' | 'setHideBackgroundGraphics'
	> &
		Partial<Pick<EditActions, 'pushRecentColor'>>,
): FormatBackgroundPanel {
	const el = createEl(doc, 'div', 'pptxv-format-background-panel');
	el.hidden = true;
	el.setAttribute('role', 'dialog');
	el.setAttribute('aria-label', t('pptx.ribbon.formatBackground'));

	const colorControl = makeColorControl(
		doc,
		{
			label: t('pptx.slideBackground.colourAriaLabel'),
			onInput: (hex) => actions.setSlideBackgroundColor(hex),
			// The committed pick (native `change`, not the drag stream) joins the
			// deck's "Recent colours" list like every other colour picker.
			onCommit: (hex) => actions.pushRecentColor?.(hex),
		},
		'#ffffff',
	);

	const clearBtn = makeButton(doc, {
		label: t('pptx.slideBackground.clearBackground'),
		text: t('pptx.slideBackground.clearBackground'),
		onClick: () => actions.clearSlideBackground(),
	});
	const closeBtn = makeButton(doc, {
		// Its own name, not the find bar's: this button was copied from
		// `find-replace-panel.ts` and announced itself as "Close find and replace".
		label: t('pptx.common.close'),
		icon: 'chevron-up',
		onClick: () => setOpen(false),
	});

	const row = createEl(doc, 'div', 'pptxv-format-background-row');
	const label = createEl(doc, 'span', 'pptxv-format-background-label');
	label.textContent = t('pptx.slideBackground.colour');
	row.append(label, colorControl.el, clearBtn.btn, closeBtn.btn);
	el.appendChild(row);

	const hideRow = createEl(doc, 'label', 'pptxv-format-background-hide-row');
	const hideCheckbox = createInspectorCheckbox(doc);
	hideCheckbox.addEventListener('change', () => {
		actions.setHideBackgroundGraphics(hideCheckbox.checked);
	});
	hideRow.append(
		hideCheckbox,
		doc.createTextNode(t('pptx.slideBackground.hideBackgroundGraphics')),
	);
	el.appendChild(hideRow);

	let open = false;
	// Where focus was when the panel opened (the ribbon command), to hand it back.
	let opener: HTMLElement | null = null;
	const setOpen = (next: boolean): void => {
		if (next === open) {
			return;
		}
		open = next;
		// Read before hiding: a browser drops focus from a control the moment its
		// container is hidden, and then the panel no longer looks focused.
		const focusInside = el.contains(doc.activeElement);
		el.hidden = !open;
		if (open) {
			const active = doc.activeElement;
			opener = active instanceof HTMLElement && !el.contains(active) ? active : null;
			// Into the panel, like the find bar, so Escape reaches it straight away.
			(el.querySelector<HTMLElement>('input:not(:disabled)') ?? closeBtn.btn).focus();
		} else {
			if (focusInside && opener?.isConnected) {
				opener.focus();
			}
			opener = null;
		}
	};
	// Escape closes the panel, as it does Svelte's (the other docked one) and
	// PowerPoint's task panes. Stopped here so the editor's own Escape (clear the
	// selection) does not also run.
	el.addEventListener('keydown', (event) => {
		if (event.key === 'Escape' && !event.isComposing) {
			event.preventDefault();
			event.stopPropagation();
			setOpen(false);
		}
	});

	return {
		el,
		toggle: () => setOpen(!open),
		setOpen,
		isOpen: () => open,
		setEditable(editable) {
			colorControl.setDisabled(!editable);
			clearBtn.setDisabled(!editable);
		},
	};
}
