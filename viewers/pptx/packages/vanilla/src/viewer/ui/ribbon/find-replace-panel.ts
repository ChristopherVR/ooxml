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
 * A docked Find & Replace panel (query + Previous / Next + replacement + the
 * four scopes + Find/Replace/Replace All), backed by the shared Find and
 * Replace session through {@link FindReplaceActions}. The status line reads
 * "n of N" for the current match; Replace acts on that match. There is no
 * in-canvas match highlighting (React's alone).
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
	const dataLabel = createEl(doc, 'label', 'pptxv-findreplace-checkbox');
	const dataInput = createInspectorCheckbox(doc);
	dataInput.setAttribute('aria-label', t('pptx.findReplace.includeChartData'));
	dataLabel.append(dataInput, doc.createTextNode(t('pptx.findReplace.includeChartData')));

	const status = createEl(doc, 'span', 'pptxv-findreplace-status');
	status.setAttribute('aria-live', 'polite');
	const hint = createEl(doc, 'p');
	hint.setAttribute('role', 'status');
	let canEdit = true;
	let open = false;
	/** The last search's count, for a host whose actions have no `status`. */
	let lastCount = 0;

	/** Paint the status line and the buttons from the session; never searches. */
	const render = (): void => {
		const query = queryInput.value;
		const current = actions.status?.() ?? { current: lastCount > 0 ? 1 : 0, total: lastCount };
		const availability = actions.replacementAvailability?.(
			query,
			matchCaseInput.checked,
			notesInput.checked,
			chartsInput.checked,
			dataInput.checked,
		);
		prevBtn.setDisabled(current.total === 0);
		nextBtn.setDisabled(current.total === 0);
		replaceBtn.setDisabled(!canEdit || !(availability?.current ?? current.total > 0));
		replaceAllBtn.setDisabled(!canEdit || !(availability?.all ?? current.total > 0));
		hint.textContent = availability?.hasReadOnly
			? t(
					dataInput.checked
						? 'pptx.findReplace.readOnlyChartText'
						: 'pptx.findReplace.readOnlyChartTitle',
				)
			: '';
		status.textContent = !open
			? ''
			: current.total > 0
				? t('pptx.findReplace.matchCount', { ...current })
				: query
					? t('pptx.findReplace.noMatches')
					: '';
	};

	const runFind = (): void => {
		lastCount = actions.search(
			queryInput.value,
			matchCaseInput.checked,
			notesInput.checked,
			chartsInput.checked,
			dataInput.checked,
		);
		render();
	};
	/** After a replace the session has searched again; a host without one is searched here. */
	const afterReplace = (): void => (actions.status ? render() : runFind());
	const step = (direction: 1 | -1): void => {
		if (direction > 0) {
			actions.next?.();
		} else {
			actions.previous?.();
		}
		render();
	};

	const prevBtn = makeButton(doc, {
		label: t('pptx.findReplace.previousMatch'),
		icon: 'chevron-up',
		onClick: () => step(-1),
	});
	const nextBtn = makeButton(doc, {
		label: t('pptx.findReplace.nextMatch'),
		icon: 'chevron-down',
		onClick: () => step(1),
	});
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
				dataInput.checked,
			);
			afterReplace();
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
				dataInput.checked,
			);
			afterReplace();
		},
	});
	const closeBtn = makeButton(doc, {
		label: t('pptx.findReplace.closeAriaLabel'),
		icon: 'chevron-up',
		onClick: () => setOpen(false),
	});

	queryInput.addEventListener('input', runFind);
	matchCaseInput.addEventListener('change', runFind);
	notesInput.addEventListener('change', runFind);
	dataInput.addEventListener('change', runFind);
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
		} else if (event.key === 'Enter') {
			event.preventDefault();
			event.stopPropagation();
			step(event.shiftKey ? -1 : 1);
		}
	});

	const row = createEl(doc, 'div', 'pptxv-findreplace-row');
	row.append(
		queryInput,
		prevBtn.btn,
		nextBtn.btn,
		replaceInput,
		matchCaseLabel,
		notesLabel,
		chartsLabel,
		dataLabel,
		findBtn.btn,
		replaceBtn.btn,
		replaceAllBtn.btn,
		closeBtn.btn,
	);
	el.append(row, status, hint);

	const setOpen = (next: boolean): void => {
		if (next === open) {
			if (open) {
				queryInput.focus();
			}
			return;
		}
		open = next;
		el.hidden = !open;
		if (!open) {
			// Closing drops the matches; the fields stay for the next open.
			actions.close?.();
			render();
			return;
		}
		queryInput.focus();
		if (queryInput.value) {
			runFind();
		} else {
			render();
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
			// Runs on every store change, so it repaints the buttons and never searches again.
			render();
		},
	};
}
