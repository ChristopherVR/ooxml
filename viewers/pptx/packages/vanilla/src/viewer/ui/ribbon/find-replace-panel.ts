import type { FindReplaceActions } from '../../editor/editor-find-replace-actions';
import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { makeButton } from '../controls';
import { createInspectorCheckbox } from '../inspector/controls-extra';

export interface FindReplacePanel {
	el: HTMLElement;
	toggle(): void;
	setOpen(open: boolean): void;
	isOpen(): boolean;
	setEditable(editable: boolean): void;
}

/**
 * A simple docked Find & Replace panel (query + replacement + match-case
 * toggle + Find/Replace/Replace All), backed by the shared `find-replace.ts`
 * helpers via {@link FindReplaceActions}. No in-canvas match highlighting or
 * next/previous cursor (see that module's docs); status text reports the
 * match count / replacement count instead.
 */
export function createFindReplacePanel(
	doc: Document,
	t: Translator,
	actions: FindReplaceActions,
): FindReplacePanel {
	const el = createEl(doc, 'div', 'pptxv-findreplace');
	el.hidden = true;
	el.setAttribute('role', 'dialog');
	el.setAttribute('aria-label', t('pptx.findReplace.ariaLabel'));

	const queryInput = doc.createElement('input');
	queryInput.type = 'text';
	queryInput.className = 'pptxv-findreplace-input';
	queryInput.placeholder = t('pptx.findReplace.findPlaceholder');
	queryInput.setAttribute('aria-label', t('pptx.findReplace.searchText'));

	const replaceInput = doc.createElement('input');
	replaceInput.type = 'text';
	replaceInput.className = 'pptxv-findreplace-input';
	replaceInput.placeholder = t('pptx.findReplace.replacePlaceholder');
	replaceInput.setAttribute('aria-label', t('pptx.findReplace.replacementText'));

	const matchCaseLabel = createEl(doc, 'label', 'pptxv-findreplace-checkbox');
	const matchCaseInput = createInspectorCheckbox(doc);
	matchCaseInput.setAttribute('aria-label', t('pptx.findReplace.toggleMatchCase'));
	matchCaseLabel.append(matchCaseInput, doc.createTextNode(t('pptx.findReplace.matchCase')));

	const notesLabel = createEl(doc, 'label', 'pptxv-findreplace-checkbox');
	const notesInput = createInspectorCheckbox(doc);
	notesInput.setAttribute('aria-label', t('pptx.findReplace.includeNotes'));
	notesLabel.append(notesInput, doc.createTextNode(t('pptx.findReplace.includeNotes')));
	const chartsLabel = createEl(doc, 'label', 'pptxv-findreplace-checkbox');
	const chartsInput = createInspectorCheckbox(doc);
	chartsInput.checked = true;
	chartsInput.setAttribute('aria-label', t('pptx.findReplace.includeChartTitles'));
	chartsLabel.append(chartsInput, doc.createTextNode(t('pptx.findReplace.includeChartTitles')));

	const status = createEl(doc, 'span', 'pptxv-findreplace-status');
	status.setAttribute('aria-live', 'polite');
	const hint = createEl(doc, 'p');
	hint.setAttribute('role', 'status');
	let canEdit = true;

	const runFind = (): void => {
		const count = actions.search(
			queryInput.value,
			matchCaseInput.checked,
			notesInput.checked,
			chartsInput.checked,
		);
		const availability = actions.replacementAvailability?.(
			queryInput.value,
			matchCaseInput.checked,
			notesInput.checked,
			chartsInput.checked,
		);
		replaceBtn.setDisabled(!canEdit || !(availability?.current ?? count > 0));
		replaceAllBtn.setDisabled(!canEdit || !(availability?.all ?? count > 0));
		hint.textContent = availability?.hasReadOnly ? t('pptx.findReplace.readOnlyChartTitle') : '';
		status.textContent =
			count > 0
				? t('pptx.findReplace.matchCount', { current: 1, total: count })
				: t('pptx.findReplace.noMatches');
	};

	const findBtn = makeButton(doc, {
		label: t('pptx.findReplace.searchText'),
		text: t('pptx.findReplace.title'),
		onClick: runFind,
	});
	const replaceBtn = makeButton(doc, {
		label: t('pptx.findReplace.replaceCurrent'),
		text: t('pptx.findReplace.replace'),
		onClick: () => {
			actions.replaceCurrent(
				queryInput.value,
				replaceInput.value,
				matchCaseInput.checked,
				notesInput.checked,
				chartsInput.checked,
			);
			runFind();
		},
	});
	const replaceAllBtn = makeButton(doc, {
		label: t('pptx.findReplace.replaceAllMatches'),
		text: t('pptx.findReplace.replaceAll'),
		onClick: () => {
			actions.replaceAll(
				queryInput.value,
				replaceInput.value,
				matchCaseInput.checked,
				notesInput.checked,
				chartsInput.checked,
			);
			runFind();
		},
	});
	const closeBtn = makeButton(doc, {
		label: t('pptx.findReplace.closeAriaLabel'),
		icon: 'chevron-up',
		onClick: () => setOpen(false),
	});

	queryInput.addEventListener('input', runFind);
	notesInput.addEventListener('change', runFind);
	chartsInput.addEventListener('change', runFind);
	// Deliberately NOT a blanket `stopPropagation()`. The shared `mapEditorKey`
	// already gates on `isEditorTextInputTarget`, and it places `find` and
	// `Escape` ABOVE those gates on purpose, so swallowing every keydown here
	// stopped Ctrl+F toggling the panel closed the way it does in the other
	// four bindings. Only keys the panel itself consumes are stopped.
	queryInput.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			event.stopPropagation();
			setOpen(false);
		}
	});

	const row = createEl(doc, 'div', 'pptxv-findreplace-row');
	row.append(
		queryInput,
		replaceInput,
		matchCaseLabel,
		notesLabel,
		chartsLabel,
		findBtn.btn,
		replaceBtn.btn,
		replaceAllBtn.btn,
		closeBtn.btn,
	);
	el.append(row, status, hint);

	let open = false;
	const setOpen = (next: boolean): void => {
		open = next;
		el.hidden = !open;
		if (open) {
			queryInput.focus();
		}
	};

	return {
		el,
		toggle: () => setOpen(!open),
		setOpen,
		isOpen: () => open,
		setEditable(editable) {
			canEdit = editable;
			replaceInput.disabled = !editable;
			replaceBtn.setDisabled(!editable);
			replaceAllBtn.setDisabled(!editable);
			if (open && queryInput.value) {
				runFind();
			}
		},
	};
}
