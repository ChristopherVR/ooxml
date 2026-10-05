import { assignKeyTips, runKeyTips, type KeyTipTarget } from '../controls';

/** Word's own KeyTips for the commands this ribbon shares with it, by the control's English label. */
const WORD_KEYS: Record<string, string> = {
	Paste: 'V',
	Cut: 'X',
	Copy: 'C',
	'Format painter': 'FP',
	'Font family': 'FF',
	'Font size': 'FS',
	'Grow font': 'FG',
	'Shrink font': 'FK',
	'Change case': 'AA',
	'Clear formatting': 'E',
	Bold: '1',
	Italic: '2',
	Underline: '3',
	Strikethrough: '4',
	Subscript: '5',
	Superscript: '6',
	'Text highlight': 'I',
	'Font color': 'FC',
	'Bulleted list': 'U',
	'Numbered list': 'N',
	'Decrease indent': 'AO',
	'Increase indent': 'AI',
	Sort: 'SO',
	'Show paragraph marks': '8',
	'Align left': 'AL',
	'Align center': 'AC',
	'Align right': 'AR',
	Justify: 'AJ',
	'Line spacing': 'K',
	Shading: 'H',
	Borders: 'B',
	'Find and replace': 'FD',
	Replace: 'R',
	'Select all': 'SL',
	Table: 'T',
	Pictures: 'P',
	Link: 'K',
	Bookmark: 'K',
	Header: 'H',
	Footer: 'O',
	'Page number': 'NU',
	Margins: 'M',
	'Page size': 'SZ',
	Orientation: 'O',
	Columns: 'J',
	'Table of contents': 'T',
	'Update table': 'U',
	Footnote: 'F',
	Endnote: 'E',
	Spelling: 'C',
	'Word count': 'W',
	'Track changes': 'G',
	Zoom: 'Q',
};

const nameOf = (el: HTMLElement) =>
	el.dataset.localearialabel ?? el.getAttribute('aria-label') ?? '';

/** The control a badge belongs on, and how to operate it. */
function targetOf(el: HTMLElement): KeyTipTarget | null {
	const wrapper = el.closest<HTMLElement>('.ribbon-menu, .ribbon-combo') ?? el;
	if (!wrapper.getClientRects().length) return null;
	const label = nameOf(el);
	if (!label) return null;
	const activate = () => {
		if (el instanceof HTMLSelectElement && !wrapper.classList.contains('ribbon-menu-gallery')) {
			el.focus();
			try {
				el.showPicker?.();
			} catch {
				// Some browsers only allow showPicker from a pointer gesture; focus is enough then.
			}
		} else if (el instanceof HTMLInputElement) el.focus();
		else wrapper.click();
	};
	return { key: WORD_KEYS[label] ?? '', element: wrapper, activate };
}

/**
 * KeyTips for the visible controls of `panel`: Word's own where this ribbon has the same command,
 * otherwise two letters from the label (see `assignKeyTips`).
 */
export function collectKeyTips(panel: HTMLElement): KeyTipTarget[] {
	const seen = new Set<HTMLElement>();
	const targets: KeyTipTarget[] = [];
	const controls = panel.querySelectorAll<HTMLElement>(
		'button[aria-label]:not(.ribbon-launcher):not([data-split-caret]):not(.ribbon-combo-caret):not(.style-tile), select[aria-label], .ribbon-combo input[aria-label], button.ribbon-overflow-button',
	);
	for (const control of controls) {
		const target = targetOf(control);
		if (!target || seen.has(target.element)) continue;
		seen.add(target.element);
		targets.push(target);
	}
	// Word's keys differ per tab, so two commands can share one; the later one is re-derived.
	assignKeyTips(targets, (target) =>
		nameOf(target.element.querySelector('[aria-label]') ?? target.element),
	);
	return targets;
}

/**
 * Shows badges over the controls of `panel` and runs the one whose letters are typed. Escape (or a
 * key with no match) ends it; `onDone` runs after any end. Returns a function that stops it.
 */
export function showCommandTips(root: HTMLElement, panel: HTMLElement, onDone: () => void) {
	return runKeyTips(root, collectKeyTips(panel), onDone, 'dve-keytip dve-keytip-command');
}
