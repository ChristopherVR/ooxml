/**
 * Collapses ribbon groups from the right, each into one dropdown button, until the panel fits;
 * the panel scrolls only when even that is not enough. Ported from docx-viewer's ribbon-overflow.ts.
 */
import { setLargeCaption } from './caption.js';
import { ribbonIcon } from './icons';

/** Icon on a collapsed group's button, by the group's English label. */
const GROUP_ICONS: Record<string, string> = {
	Clipboard: 'paste',
	Font: 'fontColor',
	Alignment: 'alignCenter',
	Number: 'numberFormat',
	Styles: 'cellStyles',
	Cells: 'insertCells',
	Editing: 'find',
	Tables: 'table',
	Illustrations: 'picture',
	Charts: 'chart',
	Links: 'link',
	Comments: 'comment',
	Text: 'textBox',
	Symbols: 'symbol',
	'Page Setup': 'margins',
	'Scale to Fit': 'zoom',
	'Sheet Options': 'gridlines',
	'Function Library': 'fx',
	'Defined Names': 'nameManager',
	'Formula Auditing': 'traceArrow',
	Calculation: 'calculate',
	'Sort & Filter': 'sortFilter',
	'Data Tools': 'dataValidation',
	Proofing: 'spelling',
	Protect: 'protect',
	'Sheet View': 'view',
	'Workbook Views': 'normalView',
	Show: 'gridlines',
	Zoom: 'zoom',
	Window: 'freeze',
};

let closeOpen: (() => void) | undefined;

function overflowButton(group: HTMLElement): HTMLButtonElement {
	const doc = group.ownerDocument;
	const button = doc.createElement('button');
	button.type = 'button';
	button.className = 'ribbon-overflow-button ribbon-large';
	button.setAttribute('aria-haspopup', 'true');
	button.append(ribbonIcon(doc, GROUP_ICONS[group.dataset.label ?? ''] ?? 'select', 28));
	button.append(doc.createElement('span'), ribbonIcon(doc, 'caret', 12));
	button.addEventListener('mousedown', (event) => event.preventDefault());
	button.addEventListener('click', () => toggleGroup(group, button));
	return button;
}

/** Shows a collapsed group's controls in a panel under its button, moving the live elements. */
function toggleGroup(group: HTMLElement, button: HTMLButtonElement): void {
	const wasOpen = button.getAttribute('aria-expanded') === 'true';
	closeOpen?.();
	if (wasOpen) return;
	const ribbon = group.closest<HTMLElement>('.xve-ribbon');
	if (!ribbon) return;
	const doc = group.ownerDocument;
	const pop = doc.createElement('div');
	pop.className = 'ribbon-popover ribbon-overflow-panel';
	const content = doc.createElement('div');
	content.className = 'ribbon-group ribbon-overflow-content';
	content.dataset.label = group.dataset.label ?? '';
	content.dataset.caption = group.dataset.caption ?? '';
	content.setAttribute('role', 'group');
	content.setAttribute('aria-label', group.getAttribute('aria-label') ?? '');
	const moved = [...group.children].filter((child) => child !== button);
	content.append(...moved);
	pop.append(content);
	ribbon.append(pop);
	const box = button.getBoundingClientRect();
	pop.style.top = `${box.bottom + 2}px`;
	const width = doc.defaultView?.innerWidth ?? 1024;
	pop.style.left = `${Math.max(4, Math.min(box.left, width - pop.offsetWidth - 8))}px`;
	button.setAttribute('aria-expanded', 'true');
	const outside = (event: Event) => {
		const path = event.composedPath();
		const inside = path.some(
			(node) => node instanceof HTMLElement && node.classList.contains('ribbon-popover'),
		);
		if (!inside && !path.includes(button)) closeOpen?.();
	};
	const escape = (event: KeyboardEvent) => {
		if (event.key !== 'Escape') return;
		closeOpen?.();
		button.focus();
	};
	pop.addEventListener('click', (event) => {
		if ((event.target as Element).closest('button[data-command]'))
			setTimeout(() => closeOpen?.(), 0);
	});
	doc.addEventListener('pointerdown', outside, true);
	doc.addEventListener('keydown', escape, true);
	closeOpen = () => {
		doc.removeEventListener('pointerdown', outside, true);
		doc.removeEventListener('keydown', escape, true);
		group.prepend(...moved);
		pop.remove();
		button.removeAttribute('aria-expanded');
		closeOpen = undefined;
	};
}

/**
 * Group captions are painted by `::after`, which does not size the group: give each group room
 * for its caption (centred, clear of the dialog launcher) so a caption never runs into the next.
 */
function sizeCaptions(panel: HTMLElement): void {
	const sizer = panel.ownerDocument.createElement('span');
	sizer.className = 'ribbon-caption-sizer';
	panel.append(sizer);
	for (const group of panel.querySelectorAll<HTMLElement>(':scope > .ribbon-group')) {
		sizer.textContent = group.dataset.caption ?? '';
		const launcher = group.querySelector(':scope > .ribbon-launcher') ? 32 : 12;
		group.style.minWidth = `${Math.ceil(sizer.offsetWidth) + launcher}px`;
	}
	sizer.remove();
}

/** Folds groups of `panel` from the right until it fits its width. */
export function fitPanel(panel: HTMLElement): void {
	if (panel.hidden || !panel.clientWidth) return;
	if (closeOpen && panel.dataset.fittedWidth === String(panel.clientWidth)) return;
	panel.dataset.fittedWidth = String(panel.clientWidth);
	closeOpen?.();
	sizeCaptions(panel);
	for (const group of panel.querySelectorAll<HTMLElement>('.ribbon-group[data-collapsed]'))
		group.removeAttribute('data-collapsed');
	const groups = [...panel.querySelectorAll<HTMLElement>(':scope > .ribbon-group')].filter(
		(group) => group.offsetWidth > 0,
	);
	for (let at = groups.length - 1; at >= 0 && panel.scrollWidth > panel.clientWidth + 1; at--) {
		const group = groups[at]!;
		let button = group.querySelector<HTMLButtonElement>(':scope > .ribbon-overflow-button');
		if (!button) {
			button = overflowButton(group);
			group.append(button);
		}
		const label = group.dataset.caption ?? group.dataset.label ?? '';
		setLargeCaption(button.querySelector('span')!, label);
		button.setAttribute('aria-label', label);
		button.title = label;
		group.setAttribute('data-collapsed', '');
	}
}

/** Keeps each panel fitted as the ribbon resizes. Returns a disposer. */
export function attachRibbonOverflow(root: HTMLElement): () => void {
	if (typeof ResizeObserver === 'undefined') return () => {};
	const observer = new ResizeObserver((entries) => {
		for (const entry of entries) fitPanel(entry.target as HTMLElement);
	});
	for (const panel of root.querySelectorAll<HTMLElement>('.ribbon-panel')) observer.observe(panel);
	return () => observer.disconnect();
}

/** Re-measures the visible panel (after a tab switch or a locale change). */
export function refitRibbon(root: HTMLElement): void {
	for (const panel of root.querySelectorAll<HTMLElement>('.ribbon-panel')) fitPanel(panel);
}
