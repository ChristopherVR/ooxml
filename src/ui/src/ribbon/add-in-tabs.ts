/**
 * Ribbon add-in tabs: tabs a host application adds after the product's own, as an Office add-in
 * does (Visio's and Excel's ACROBAT tab). The host describes each tab with plain data; every
 * product editor exposes the list as its `ribbonAddIns` property and renders it with
 * `syncRibbonAddIns`, so one descriptor works in every editor. An add-in command never reaches
 * the product's own command router: it runs its `run` callback and is announced with the
 * `office-ribbon-add-in` event, which leaves the editor's shadow tree. The descriptor itself is
 * DOM-free and lives in `ooxml-core/ribbon`, so the binding contracts in the core share it.
 */

import {
	acceptedRibbonAddIns,
	findRibbonAddInCommand,
	RIBBON_ADD_IN_EVENT,
	type RibbonAddInCommand,
	type RibbonAddInCommandDetail,
	type RibbonAddInGroup,
	type RibbonAddInTab,
} from 'ooxml-core/ribbon';

export {
	RIBBON_ADD_IN_EVENT,
	type RibbonAddInCommand,
	type RibbonAddInCommandDetail,
	type RibbonAddInGroup,
	type RibbonAddInTab,
} from 'ooxml-core/ribbon';

/** Bubbling and composed, dispatched from the ribbon when an add-in command is chosen. */
export type OfficeRibbonAddInEvent = CustomEvent<RibbonAddInCommandDetail>;

export interface RibbonAddInOptions {
	/** Class of each tab panel, for products that style their own panels (`ribbon-content`). */
	panelClass?: string;
	/** Wraps the groups of a panel, for products whose panels hold one toolbar row. */
	wrap?: (doc: Document, tab: RibbonAddInTab, groups: HTMLElement[]) => HTMLElement[];
}

const SMALL_ROWS = 3;

function button(doc: Document, spec: RibbonAddInCommand): HTMLElement {
	const menu = !!spec.items?.length;
	const el = doc.createElement(menu ? 'office-ui-menu-button' : 'office-ui-button');
	if (!menu) el.setAttribute('command', spec.id);
	el.dataset.addInCommand = spec.id;
	el.setAttribute('label', spec.label);
	el.setAttribute('title', spec.title ?? spec.label);
	if (spec.icon) el.setAttribute('icon', spec.icon);
	if (spec.size !== 'small') el.setAttribute('variant', 'stacked');
	if (spec.disabled) el.setAttribute('disabled', '');
	for (const item of spec.items ?? []) {
		const entry = doc.createElement('office-ui-menu-item');
		entry.setAttribute('command', item.id);
		entry.dataset.addInCommand = item.id;
		entry.setAttribute('label', item.label);
		if (item.title) entry.setAttribute('title', item.title);
		if (item.icon) entry.setAttribute('icon', item.icon);
		if (item.disabled) entry.setAttribute('disabled', '');
		el.append(entry);
	}
	return el;
}

function group(doc: Document, spec: RibbonAddInGroup): HTMLElement {
	const el = doc.createElement('office-ui-ribbon-group');
	el.setAttribute('label', spec.label);
	let column: HTMLElement | undefined;
	for (const command of spec.commands) {
		if (command.size !== 'small') {
			column = undefined;
			el.append(button(doc, command));
			continue;
		}
		if (!column || column.childElementCount >= SMALL_ROWS) {
			column = doc.createElement('office-ui-ribbon-stack');
			el.append(column);
		}
		column.append(button(doc, command));
	}
	return el;
}

/** The tab ids a product ribbon already uses for panels that are not add-in tabs. */
const ownTabs = (ribbon: HTMLElement): Set<string> =>
	new Set(
		[...ribbon.children].flatMap((child) =>
			child instanceof HTMLElement && child.dataset.ribbonTab && !('addIn' in child.dataset)
				? [child.dataset.ribbonTab]
				: [],
		),
	);

/**
 * Replaces the add-in panels of `ribbon` (an `office-ui-ribbon`) with `tabs`, after the product's
 * own panels and in the given order. A tab whose id is empty, repeated or already a product tab
 * is skipped, so an add-in can never replace a built-in tab. The selected tab is kept when it
 * still exists. Returns the ids that were added.
 */
export function syncRibbonAddIns(
	ribbon: HTMLElement,
	tabs: readonly RibbonAddInTab[],
	options: RibbonAddInOptions = {},
): string[] {
	const doc = ribbon.ownerDocument;
	const accepted = acceptedRibbonAddIns(tabs, ownTabs(ribbon));
	for (const child of [...ribbon.children])
		if (child instanceof HTMLElement && 'addIn' in child.dataset) child.remove();
	for (const tab of accepted) {
		const panel = doc.createElement('div');
		if (options.panelClass) panel.className = options.panelClass;
		panel.id = `add-in-${tab.id}-panel`;
		panel.dataset.addIn = '';
		panel.dataset.ribbonTab = tab.id;
		panel.dataset.label = tab.label;
		if (tab.keytip) panel.dataset.tabKeytip = tab.keytip;
		const groups = tab.groups.map((spec) => group(doc, spec));
		panel.append(...(options.wrap ? options.wrap(doc, tab, groups) : groups));
		panel.addEventListener('office-command', (event) => {
			// The product's router never sees an add-in command.
			event.stopPropagation();
			const id = (event as CustomEvent<{ command?: string }>).detail?.command;
			const command = findRibbonAddInCommand(tab, id);
			if (!command) return;
			command.run?.();
			ribbon.dispatchEvent(
				new CustomEvent(RIBBON_ADD_IN_EVENT, {
					detail: { tab: tab.id, command: command.id },
					bubbles: true,
					composed: true,
				}) satisfies OfficeRibbonAddInEvent,
			);
		});
		ribbon.append(panel);
	}
	return accepted.map((tab) => tab.id);
}
