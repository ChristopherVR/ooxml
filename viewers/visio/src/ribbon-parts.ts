import { emitRibbonAction, type VisioRibbonAction } from './ribbon-action.js';

/** A shared `office-ui-button` with Visio-specific state hooks. */
export type RibbonCommand = HTMLElement & { disabled: boolean };

/**
 * A large stacked ribbon command. `id` is a stable name for state sync and tests; the typed
 * `action` is what runs. Labels, icons and shortcuts are static application strings.
 */
export function tool(
	doc: Document,
	id: string,
	label: string,
	icon: string,
	action: VisioRibbonAction,
	options: { keys?: [aria: string, hint: string]; pressed?: boolean; controls?: string } = {},
): RibbonCommand {
	const el = doc.createElement('office-ui-button') as RibbonCommand;
	el.setAttribute('variant', 'stacked');
	el.setAttribute('command', id);
	el.setAttribute('icon', icon);
	el.setAttribute('label', label);
	el.setAttribute('title', options.keys ? `${label} (${options.keys[1]})` : label);
	if (options.keys) el.setAttribute('keyshortcuts', options.keys[0]);
	if (options.pressed !== undefined) el.setAttribute('pressed', String(options.pressed));
	if (options.controls) el.setAttribute('aria-controls', options.controls);
	el.addEventListener('office-command', (event) => {
		event.stopPropagation();
		emitRibbonAction(el, action);
	});
	return el;
}

/** A labelled ribbon group (the shared `office-ui-ribbon-group`). */
export function group(
	doc: Document,
	label: string,
	children: readonly HTMLElement[],
	className = '',
): HTMLElement {
	const el = doc.createElement('office-ui-ribbon-group');
	el.setAttribute('label', label);
	if (className) el.className = className;
	el.append(...children);
	return el;
}

/** A keyboard-navigable row of groups (the shared `office-ui-toolbar`). */
export function commandRow(doc: Document, label: string, children: readonly HTMLElement[]) {
	const el = doc.createElement('office-ui-toolbar');
	el.className = 'ribbon-row';
	el.setAttribute('aria-label', label);
	el.append(...children);
	return el;
}

/** Muted two-line guidance at the end of a ribbon panel. */
export function hint(doc: Document, lines: readonly string[]): HTMLElement {
	const el = doc.createElement('p');
	el.className = 'ribbon-hint';
	lines.forEach((line, index) => {
		if (index) el.append(doc.createElement('br'));
		el.append(line);
	});
	return el;
}
