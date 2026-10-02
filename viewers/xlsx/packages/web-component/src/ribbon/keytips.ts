/**
 * KeyTips (Alt or F10): badges over the tabs, then over the chosen tab's commands; typing the
 * letters runs the control. Excel's own keys where this ribbon has the same command, otherwise
 * two letters from the label, never the prefix of another. Ported from docx-viewer's
 * ribbon-keytips.ts and ribbon-behavior.ts.
 */

/** Excel's tab KeyTips, by tab id (File is `F`). */
const TAB_KEYS: Record<string, string> = {
	home: 'H',
	insert: 'N',
	draw: 'JI',
	'page-layout': 'P',
	pageLayout: 'P',
	formulas: 'M',
	data: 'A',
	review: 'R',
	view: 'W',
	help: 'Y',
	'table-design': 'JT',
	tableDesign: 'JT',
	'chart-design': 'JC',
	chartDesign: 'JC',
};

/** Excel's command KeyTips, by the command's English label. */
const COMMAND_KEYS: Record<string, string> = {
	Paste: 'V',
	Cut: 'X',
	Copy: 'C',
	'Format Painter': 'FP',
	Font: 'FF',
	'Font Size': 'FS',
	'Increase Font Size': 'FG',
	'Decrease Font Size': 'FK',
	Bold: '1',
	Italic: '2',
	Underline: '3',
	Borders: 'B',
	'Fill Color': 'H',
	'Font Color': 'FC',
	'Top Align': 'AT',
	'Middle Align': 'AM',
	'Bottom Align': 'AB',
	'Align Left': 'AL',
	Center: 'AC',
	'Align Right': 'AR',
	'Wrap Text': 'W',
	'Merge & Center': 'M',
	'Number Format': 'N',
	'Accounting Number Format': 'AN',
	'Percent Style': 'P',
	'Comma Style': 'K',
	'Increase Decimal': '0',
	'Decrease Decimal': '9',
	'Conditional Formatting': 'L',
	'Format as Table': 'T',
	'Cell Styles': 'J',
	Insert: 'I',
	Delete: 'D',
	Format: 'O',
	AutoSum: 'U',
	Fill: 'FI',
	Clear: 'E',
	'Sort & Filter': 'S',
	'Find & Select': 'FD',
};

interface Target {
	key: string;
	element: HTMLElement;
	activate(): void;
}

const labelOf = (node: HTMLElement) =>
	node.dataset.labelKey ?? node.getAttribute('aria-label') ?? '';

function badge(doc: Document, key: string, className: string): HTMLSpanElement {
	const node = doc.createElement('span');
	node.className = className;
	node.textContent = key;
	node.setAttribute('aria-hidden', 'true');
	return node;
}

/** Tips for the visible controls of `panel`. */
export function collectKeyTips(panel: HTMLElement): Target[] {
	const targets: Target[] = [];
	const seen = new Set<HTMLElement>();
	for (const control of panel.querySelectorAll<HTMLElement>(
		'button[data-command]:not([data-launcher]), select[data-command], input[data-command], .ribbon-menu-button, .ribbon-overflow-button',
	)) {
		if (
			seen.has(control) ||
			!control.getClientRects().length ||
			control.closest('[data-xve-hidden]')
		)
			continue;
		seen.add(control);
		const activate = () => {
			if (control instanceof HTMLSelectElement || control instanceof HTMLInputElement) {
				control.focus();
				try {
					if (control instanceof HTMLSelectElement) control.showPicker?.();
				} catch {
					// showPicker needs a user gesture in some browsers; focus is enough then.
				}
			} else control.click();
		};
		targets.push({ key: COMMAND_KEYS[labelOf(control)] ?? '', element: control, activate });
	}
	const used: string[] = [];
	const clash = (key: string) => used.some((u) => u.startsWith(key) || key.startsWith(u));
	for (const target of targets) {
		if (target.key && !clash(target.key)) used.push(target.key);
		else target.key = '';
	}
	for (const target of targets.filter((item) => !item.key)) {
		const source =
			(target.element.getAttribute('aria-label') ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '') ||
			'ZZ';
		let key = '';
		for (let a = 0; a < source.length && !key; a++)
			for (let b = a + 1; b < source.length && !key; b++) {
				const candidate = source[a]! + source[b]!;
				if (!clash(candidate)) key = candidate;
			}
		for (let n = 10; !key; n++) if (!clash(String(n))) key = String(n);
		used.push(key);
		target.key = key;
	}
	return targets;
}

/**
 * Shows badges over `targets` and runs the one whose letters are typed; Escape, a pointer press
 * or a key with no match ends it. Returns a function that stops it.
 */
function runTips(
	root: HTMLElement,
	targets: Target[],
	fixed: boolean,
	onDone: () => void,
): () => void {
	const doc = root.ownerDocument;
	const shown = targets.map((target) => {
		const node = badge(doc, target.key, fixed ? 'xve-keytip xve-keytip-command' : 'xve-keytip');
		if (fixed) {
			const box = target.element.getBoundingClientRect();
			node.style.left = `${box.left + box.width / 2}px`;
			node.style.top = `${box.bottom - 10}px`;
			root.append(node);
		} else target.element.append(node);
		return { node, target };
	});
	let typed = '';
	const stop = () => {
		for (const { node } of shown) node.remove();
		root.removeEventListener('keydown', onKey, true);
		root.removeEventListener('pointerdown', stop, true);
		onDone();
	};
	function onKey(event: KeyboardEvent) {
		if (event.key === 'Alt' || event.key === 'Shift') return;
		event.preventDefault();
		event.stopPropagation();
		if (event.key === 'Escape' || event.ctrlKey || event.metaKey) return stop();
		typed += event.key.toUpperCase();
		const matches = shown.filter(({ target }) => target.key.startsWith(typed));
		if (!matches.length) return stop();
		for (const { node, target } of shown) node.hidden = !target.key.startsWith(typed);
		const exact = matches.find(({ target }) => target.key === typed);
		if (exact) {
			stop();
			exact.target.activate();
		}
	}
	root.addEventListener('keydown', onKey, true);
	root.addEventListener('pointerdown', stop, true);
	return stop;
}

let active: (() => void) | undefined;

/** Level one: tab badges; picking a tab opens it and shows its command badges (level two). */
export function showTabKeyTips(
	root: HTMLElement,
	tabsBar: HTMLElement,
	fileTab: HTMLElement,
	select: (id: string) => void,
): void {
	active?.();
	const targets: Target[] = [{ key: 'F', element: fileTab, activate: () => fileTab.click() }];
	for (const tab of tabsBar.querySelectorAll<HTMLElement>('[role="tab"]')) {
		const id = tab.dataset.tab ?? '';
		if (tab.hidden || tab.hasAttribute('data-xve-hidden')) continue;
		const used = targets.map((target) => target.key);
		let key = TAB_KEYS[id] ?? '';
		if (!key || used.includes(key))
			key =
				[...(tab.textContent ?? id).toUpperCase().replace(/[^A-Z]/g, '')].find(
					(ch) => !used.includes(ch),
				) ?? `Z${used.length}`;
		targets.push({
			key,
			element: tab,
			activate: () => {
				select(id);
				tab.focus();
				const panel = root.querySelector<HTMLElement>(
					`#${tab.getAttribute('aria-controls') ?? ''}`,
				);
				if (panel) active = runTips(root, collectKeyTips(panel), true, () => (active = undefined));
			},
		});
	}
	active = runTips(root, targets, false, () => (active = undefined));
}

export function hideKeyTips(): void {
	active?.();
}
