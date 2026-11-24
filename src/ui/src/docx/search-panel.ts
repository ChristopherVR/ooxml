import type { EditorView } from 'prosemirror-view';
import type { FindFormatCriteria, ReplaceFormatPatch } from 'ooxml-core/docx';
import {
	SearchController,
	type SearchOptions,
	type SearchStatus,
	type SearchStoryHost,
} from './search-controller';
import { createFormatPicker } from './search-format-controls';
import {
	OPTION_LABELS,
	SPECIAL_CODES,
	storyLabel,
	WILDCARD_DISABLED,
} from './search-panel-options';
import {
	localizeElement,
	normalizeEditorLocale,
	translate,
	type EditorLocale,
} from './localization';

export interface SearchPanelOptions {
	getView: () => EditorView | undefined;
	/** Headers, footers, notes and text boxes; without it Find searches the body only. */
	stories?: SearchStoryHost;
	onClose?: () => void;
}

export interface SearchPanelHandle {
	element: HTMLElement;
	open(focus?: 'find' | 'replace'): void;
	close(): void;
	refresh(): void;
	setLocale(locale: string): void;
	readonly isOpen: boolean;
}

const styleText = `
	.dve-search-panel{display:flex;align-items:end;flex-wrap:wrap;gap:8px;padding:9px 12px;border-bottom:1px solid var(--line,#ddd);background:var(--surface,#fff);color:var(--ink,#222);font:12px/1.35 'Segoe UI',Arial,sans-serif}
	.dve-search-panel[hidden]{display:none}
	.dve-search-panel label{display:flex;flex-direction:column;gap:3px;color:var(--muted,#666)}
	.dve-search-panel input[type=search],.dve-search-panel input[type=text]{box-sizing:border-box;width:clamp(130px,20vw,220px);height:29px;padding:4px 7px;border:1px solid var(--line,#ddd);border-radius:3px;background:var(--surface,#fff);color:var(--ink,#222);font:inherit}
	.dve-search-panel .case{flex-direction:row;align-items:center;align-self:center;gap:5px;white-space:nowrap}
	.dve-search-panel .dve-search-options{display:flex;flex-basis:100%;flex-wrap:wrap;gap:6px 16px;padding-top:4px}
	.dve-search-panel .dve-search-options[hidden]{display:none}
	.dve-search-panel .dve-search-options label{flex-direction:row;align-items:center;gap:5px;white-space:nowrap}
	.dve-search-panel .dve-search-options label[aria-disabled=true]{opacity:.55}
	.dve-search-panel select{height:29px;padding:0 6px;border:1px solid var(--line,#ddd);border-radius:3px;background:var(--surface,#fff);color:var(--ink,#222);font:inherit}
	.dve-search-panel button{height:29px;padding:0 9px;border:1px solid var(--line,#ddd);border-radius:3px;background:var(--surface,#fff);color:var(--ink,#222);font:inherit;cursor:pointer}
	.dve-search-panel button:hover:not(:disabled){border-color:var(--blue,#1769aa)}
	.dve-search-panel button:disabled{opacity:.55;cursor:default}
	.dve-search-panel [role=status]{min-width:100px;color:var(--muted,#666)}
	.dve-search-panel .dve-search-story{color:var(--muted,#666);font-style:italic}
	.dve-search-panel .dve-search-format{display:flex;flex-basis:100%;flex-wrap:wrap;align-items:center;gap:6px;margin:4px 0 0;padding:6px 8px;border:1px solid var(--line,#ddd);border-radius:4px}
	.dve-search-panel .dve-search-format legend{padding:0 4px;color:var(--muted,#666);font-size:11px}
	.dve-search-panel .dve-search-size{width:62px;height:29px;padding:4px 6px;border:1px solid var(--line,#ddd);border-radius:3px;background:var(--surface,#fff);color:var(--ink,#222);font:inherit}
	.dve-search-panel .dve-search-color{width:84px;height:29px;padding:4px 6px;border:1px solid var(--line,#ddd);border-radius:3px;background:var(--surface,#fff);color:var(--ink,#222);font:inherit}
	.dve-search-panel .dve-search-format-summary{color:var(--muted,#666);font-size:11px;font-style:italic}
`;

export function createSearchPanel(
	options: SearchPanelOptions,
	initialLocale = 'en',
): SearchPanelHandle {
	let locale: EditorLocale = normalizeEditorLocale(initialLocale);
	const controller = new SearchController(options.getView, options.stories);
	const panel = document.createElement('section');
	panel.className = 'dve-search-panel';
	panel.setAttribute('role', 'region');
	panel.setAttribute('aria-label', 'Find and replace');
	panel.hidden = true;
	const style = document.createElement('style');
	style.textContent = styleText;
	const findLabel = document.createElement('label');
	findLabel.textContent = 'Find text';
	const findInput = document.createElement('input');
	findInput.type = 'search';
	findInput.setAttribute('aria-label', 'Find text');
	findLabel.append(findInput);
	const replaceLabel = document.createElement('label');
	replaceLabel.textContent = 'Replace with';
	const replaceInput = document.createElement('input');
	replaceInput.type = 'text';
	replaceInput.setAttribute('aria-label', 'Replace with');
	replaceLabel.append(replaceInput);
	const caseLabel = document.createElement('label');
	caseLabel.className = 'case';
	const caseInput = document.createElement('input');
	caseInput.type = 'checkbox';
	caseInput.setAttribute('aria-label', 'Match case');
	caseLabel.append(caseInput, document.createTextNode('Match case'));
	const button = (label: string) => {
		const el = document.createElement('button');
		el.type = 'button';
		el.textContent = label;
		el.setAttribute('aria-label', label);
		return el;
	};
	const previous = button('Find previous');
	const next = button('Find next');
	const replace = button('Replace');
	const replaceAll = button('Replace all');
	const more = button('More >>');
	more.setAttribute('aria-expanded', 'false');
	const close = button('Close search');
	const result = document.createElement('span');
	result.setAttribute('role', 'status');
	result.setAttribute('aria-label', 'Search results');
	result.setAttribute('aria-live', 'polite');
	const where = document.createElement('span');
	where.className = 'dve-search-story';
	where.setAttribute('aria-label', 'Search location');

	// More >> options: each checkbox by its option name, so the wildcard rule can disable them.
	const optionsSection = document.createElement('div');
	optionsSection.className = 'dve-search-options';
	optionsSection.setAttribute('role', 'group');
	optionsSection.setAttribute('aria-label', 'Search options');
	optionsSection.hidden = true;
	const optionBoxes = new Map<string, { box: HTMLInputElement; label: HTMLLabelElement }>();
	for (const [name, text] of OPTION_LABELS) {
		const label = document.createElement('label');
		const box = document.createElement('input');
		box.type = 'checkbox';
		box.setAttribute('aria-label', text);
		label.append(box, document.createTextNode(text));
		optionsSection.append(label);
		optionBoxes.set(name, { box, label });
	}

	// The Special menu inserts a Find code at the cursor of the field that last had focus.
	let lastField: HTMLInputElement = findInput;
	const special = document.createElement('select');
	special.className = 'dve-search-special';
	special.setAttribute('aria-label', 'Special characters');
	const placeholder = document.createElement('option');
	placeholder.value = '';
	placeholder.textContent = 'Special characters';
	special.append(placeholder);
	for (const [label, code] of SPECIAL_CODES) {
		const option = document.createElement('option');
		option.value = code;
		option.textContent = label;
		special.append(option);
	}
	const insertCode = (code: string) => {
		const start = lastField.selectionStart ?? lastField.value.length;
		const end = lastField.selectionEnd ?? start;
		lastField.value = lastField.value.slice(0, start) + code + lastField.value.slice(end);
		lastField.setSelectionRange(start + code.length, start + code.length);
		lastField.dispatchEvent(new Event('input', { bubbles: true }));
		lastField.focus();
	};
	special.addEventListener('change', () => {
		if (special.value) insertCode(special.value);
		special.value = '';
	});
	for (const field of [findInput, replaceInput])
		field.addEventListener('focus', () => {
			lastField = field;
		});

	const styleNames = () => Object.keys(options.stories?.model().paragraphStyles?.styles ?? {});
	const findFormat = createFormatPicker<FindFormatCriteria>(
		'find',
		() => {
			findFormat.refreshSummary(locale);
			refreshQuery();
		},
		styleNames,
	);
	const replaceFormat = createFormatPicker<ReplaceFormatPatch>(
		'replace',
		() => {
			replaceFormat.refreshSummary(locale);
		},
		styleNames,
	);
	optionsSection.append(findFormat.element, replaceFormat.element);

	const checked = (name: string): boolean => optionBoxes.get(name)?.box.checked === true;
	const currentOptions = (): SearchOptions => {
		const format = findFormat.value();
		return {
			matchCase: caseInput.checked,
			wholeWords: checked('wholeWords'),
			wildcards: checked('wildcards'),
			matchPrefix: checked('matchPrefix'),
			matchSuffix: checked('matchSuffix'),
			ignorePunctuation: checked('ignorePunctuation'),
			ignoreWhitespace: checked('ignoreWhitespace'),
			...(format ? { format } : {}),
		};
	};
	const syncWildcards = () => {
		const wildcards = optionBoxes.get('wildcards')?.box.checked ?? false;
		for (const name of WILDCARD_DISABLED) {
			const entry = optionBoxes.get(name);
			if (!entry) continue;
			entry.box.disabled = wildcards;
			entry.label.setAttribute('aria-disabled', String(wildcards));
		}
	};

	const renderStatus = (status: SearchStatus) => {
		result.textContent = status.invalid
			? translate(locale, 'Invalid search pattern')
			: !status.query && !status.hasFormat
				? translate(locale, 'Enter text to search')
				: !status.count
					? translate(locale, 'No matches')
					: status.active
						? `${status.active} ${translate(locale, 'of')} ${status.count} ${translate(locale, status.count === 1 ? 'matches one' : 'matches many')}`
						: `${status.count} ${translate(locale, status.count === 1 ? 'matches one' : 'matches many')}`;
		where.textContent = options.stories && status.active ? storyLabel(locale, status.story) : '';
		replace.disabled = status.readOnly || !status.count;
		replaceAll.disabled = status.readOnly || !status.count;
		previous.disabled = next.disabled = !status.count;
	};
	const refreshQuery = () => renderStatus(controller.search(findInput.value, currentOptions()));
	findInput.addEventListener('input', refreshQuery);
	caseInput.addEventListener('change', refreshQuery);
	for (const [name] of OPTION_LABELS)
		optionBoxes.get(name)?.box.addEventListener('change', () => {
			syncWildcards();
			refreshQuery();
		});
	more.addEventListener('click', () => {
		const expanded = optionsSection.hidden;
		optionsSection.hidden = !expanded;
		more.setAttribute('aria-expanded', String(expanded));
		more.textContent = translate(locale, expanded ? '<< Less' : 'More >>');
		more.setAttribute('aria-label', more.textContent);
	});
	// Opening a header or note editor focuses it; Find keeps the focus in the panel, as Word's does.
	const run = (action: () => unknown) => () => {
		const root = panel.getRootNode() as Document | ShadowRoot;
		const focused = root.activeElement;
		action();
		if (focused instanceof HTMLElement && panel.contains(focused) && root.activeElement !== focused)
			focused.focus();
		renderStatus(controller.status);
	};
	previous.addEventListener(
		'click',
		run(() => controller.findPrevious()),
	);
	next.addEventListener(
		'click',
		run(() => controller.findNext()),
	);
	replace.addEventListener(
		'click',
		run(() => controller.replaceOne(replaceInput.value, replaceFormat.value())),
	);
	replaceAll.addEventListener(
		'click',
		run(() => controller.replaceAll(replaceInput.value, replaceFormat.value())),
	);
	findInput.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			if (event.shiftKey) previous.click();
			else next.click();
		}
	});
	panel.addEventListener('keydown', (event) => {
		if (event.key === 'Escape' && !panel.hidden) {
			event.preventDefault();
			closePanel();
		}
	});

	let isOpen = false;
	const closePanel = () => {
		isOpen = false;
		panel.hidden = true;
		options.onClose?.();
	};
	close.addEventListener('click', closePanel);
	panel.append(
		style,
		findLabel,
		replaceLabel,
		caseLabel,
		special,
		previous,
		next,
		replace,
		replaceAll,
		more,
		result,
		where,
		close,
		optionsSection,
	);
	localizeElement(panel, locale);
	renderStatus(controller.status);
	return {
		element: panel,
		get isOpen() {
			return isOpen;
		},
		open(focus = 'find') {
			isOpen = true;
			panel.hidden = false;
			this.refresh();
			(focus === 'replace' ? replaceInput : findInput).focus();
		},
		close: closePanel,
		refresh() {
			// The editor refreshes on every change; a closed panel has nothing to show.
			if (!isOpen) return;
			controller.search(findInput.value, currentOptions());
			renderStatus(controller.status);
		},
		setLocale(value: string) {
			locale = normalizeEditorLocale(value);
			localizeElement(panel, locale);
			findFormat.refreshSummary(locale);
			replaceFormat.refreshSummary(locale);
			more.textContent = translate(locale, optionsSection.hidden ? 'More >>' : '<< Less');
			renderStatus(controller.status);
		},
	};
}
