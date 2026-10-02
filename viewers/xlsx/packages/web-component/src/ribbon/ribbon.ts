/**
 * The ribbon renderer: the File tab, the registered tabs (contextual ones only while their test
 * passes), one panel per tab with labelled groups and dialog launchers, collapse/peek, KeyTips,
 * hidden actions and overflow folding. Structure and look follow docx-viewer's ribbon.ts.
 */
import type { EditorContext } from '../context';
import { commandButton, el, renderControl, type RenderScope } from './controls';
import { ribbonIcon } from './icons';
import { attachRibbonOverflow, refitRibbon } from './overflow';
import {
	onRibbonTabsChange,
	ribbonTabs,
	type RibbonControl,
	type RibbonGroup,
	type RibbonTab,
} from './parts';
import { closeRibbonPopover } from './popover';
import { showTabKeyTips } from './keytips';
import { applyRibbonVisibility } from './visibility';

export interface RibbonHandlers {
	openBackstage(): void;
	isHidden(id: string): boolean;
}

export interface Ribbon {
	readonly element: HTMLElement;
	/** Re-reads every control state (enabled, pressed, values) and the contextual tabs. */
	refresh(): void;
	/** Rebuilds all tabs (after a locale change or newly registered tabs). */
	rebuild(): void;
	selectTab(id: string): void;
	activeTab(): string | undefined;
	/** Alt / F10: focus the selected tab and show the tab KeyTips. */
	showKeyTips(): void;
	focus(): void;
	destroy(): void;
}

const isSmall = (control: RibbonControl): boolean =>
	control.kind === 'button' ||
	control.kind === 'toggle' ||
	control.kind === 'split' ||
	control.kind === 'menu'
		? control.size !== 'large'
		: control.kind === 'color' || control.kind === 'select';

/**
 * Lays out a run of small controls the way Excel's groups look: separators break rows (Font,
 * Alignment), a leading drop-down gets a row of its own (Number), anything else fills columns of
 * three (Clipboard's Cut/Copy/Format Painter, Editing's AutoSum/Fill/Clear).
 */
function layoutSmall(scope: RenderScope, run: RibbonControl[]): HTMLElement[] {
	const { doc } = scope;
	const render = (controls: RibbonControl[]) => {
		const box = el(doc, 'div', 'ribbon-stack');
		for (const control of controls) {
			const node = renderControl(scope, control, true);
			if (node) box.append(node);
		}
		return box;
	};
	let lines: RibbonControl[][] | undefined;
	if (run.some((control) => control.kind === 'separator')) {
		lines = [[]];
		for (const control of run)
			if (control.kind === 'separator') lines.push([]);
			else lines[lines.length - 1]!.push(control);
	} else if (run[0]?.kind === 'select' && run.some((control) => control.kind !== 'select')) {
		const lead = run.findIndex((control) => control.kind !== 'select');
		lines = [run.slice(0, lead), run.slice(lead)];
	}
	if (lines) {
		const column = el(doc, 'div', 'ribbon-stack');
		for (const line of lines.filter((items) => items.length)) {
			const row = render(line);
			if (row.childElementCount) column.append(row);
		}
		return column.childElementCount ? [column] : [];
	}
	const columns: HTMLElement[] = [];
	for (let at = 0; at < run.length; at += 3) {
		const box = render(run.slice(at, at + 3));
		if (box.childElementCount) columns.push(box);
	}
	return columns;
}

function renderGroup(scope: RenderScope, group: RibbonGroup): HTMLElement | null {
	const { ctx, doc } = scope;
	const node = el(doc, 'div', 'ribbon-group');
	const caption = ctx.t(group.label);
	node.dataset.group = group.id;
	node.dataset.label = group.label;
	node.dataset.caption = caption;
	node.setAttribute('role', 'group');
	node.setAttribute('aria-label', caption);
	let run: RibbonControl[] = [];
	const flush = () => {
		node.append(...layoutSmall(scope, run));
		run = [];
	};
	for (const control of group.controls) {
		if (isSmall(control) || control.kind === 'separator') {
			run.push(control);
			continue;
		}
		flush();
		const rendered = renderControl(scope, control, false);
		if (rendered) node.append(rendered);
	}
	flush();
	const launcher = group.launcher ? ctx.commands.get(group.launcher) : undefined;
	if (launcher) {
		const button = commandButton(scope, launcher, { size: 'small' });
		button.className = 'ribbon-launcher';
		button.replaceChildren(ribbonIcon(doc, 'launcher', 10));
		button.dataset.launcher = '';
		node.append(button);
	}
	return node.childElementCount ? node : null;
}

export function createRibbon(ctx: EditorContext, handlers: RibbonHandlers): Ribbon {
	const doc = ctx.host.ownerDocument;
	const root = el(doc, 'div', 'xve-ribbon');
	root.setAttribute('part', 'ribbon');
	root.setAttribute('role', 'toolbar');
	const tabsBar = el(doc, 'nav', 'ribbon-tabs');
	tabsBar.setAttribute('role', 'tablist');
	let scope: RenderScope = { ctx, doc, updates: [], isHidden: handlers.isHidden };
	let tabs: RibbonTab[] = [];
	let selected = 'home';
	let stopOverflow = () => {};

	const fileTab = el(doc, 'button', 'xve-file-tab');
	fileTab.type = 'button';
	fileTab.setAttribute('aria-haspopup', 'dialog');
	fileTab.addEventListener('click', () => handlers.openBackstage());
	const collapse = el(doc, 'button', 'ribbon-collapse');
	collapse.type = 'button';
	collapse.setAttribute('aria-pressed', 'false');
	collapse.append(ribbonIcon(doc, 'previous', 14));

	const peek = (on: boolean) => root.toggleAttribute('data-peek', on);
	const setCollapsed = (value: boolean) => {
		root.toggleAttribute('data-collapsed', value);
		collapse.setAttribute('aria-pressed', String(value));
		peek(false);
	};
	collapse.addEventListener('click', () => setCollapsed(!root.hasAttribute('data-collapsed')));
	root.addEventListener('keydown', (event) => {
		if (event.key === 'F1' && event.ctrlKey) {
			event.preventDefault();
			setCollapsed(!root.hasAttribute('data-collapsed'));
		} else if (event.key === 'Escape' && root.hasAttribute('data-peek')) peek(false);
	});
	root.addEventListener('click', (event) => {
		if ((event.target as Element).closest?.('.ribbon-panel button[data-command]')) peek(false);
	});

	const tabButtons = () => [...tabsBar.querySelectorAll<HTMLButtonElement>('[role="tab"]')];
	const panels = () => [...root.querySelectorAll<HTMLElement>('.ribbon-panel')];
	const visibleTabs = () =>
		tabButtons().filter((tab) => !tab.hidden && !tab.hasAttribute('data-xve-hidden'));

	const select = (id: string) => {
		const target = tabButtons().find((tab) => tab.dataset.tab === id);
		if (!target || target.hidden) return;
		selected = id;
		closeRibbonPopover();
		for (const tab of tabButtons()) {
			const on = tab === target;
			tab.setAttribute('aria-selected', String(on));
			tab.tabIndex = on ? 0 : -1;
		}
		for (const panel of panels()) panel.hidden = panel.dataset.tab !== id;
		if (root.hasAttribute('data-collapsed')) peek(true);
		refitRibbon(root);
	};

	const build = () => {
		stopOverflow();
		closeRibbonPopover();
		scope = { ctx, doc, updates: [], isHidden: handlers.isHidden };
		tabs = [...ribbonTabs()];
		root.setAttribute('aria-label', ctx.t('Ribbon'));
		tabsBar.setAttribute('aria-label', ctx.t('Ribbon tabs'));
		fileTab.textContent = ctx.t('File');
		collapse.setAttribute('aria-label', ctx.t('Collapse the ribbon'));
		collapse.title = collapse.getAttribute('aria-label') ?? '';
		tabsBar.replaceChildren(fileTab);
		for (const panel of panels()) panel.remove();
		for (const tab of tabs) {
			const button = el(doc, 'button');
			button.type = 'button';
			button.id = `xve-tab-${tab.id}`;
			button.dataset.tab = tab.id;
			button.textContent = ctx.t(tab.label);
			button.setAttribute('role', 'tab');
			button.setAttribute('aria-controls', `xve-panel-${tab.id}`);
			if (tab.contextual) button.dataset.contextual = '';
			button.addEventListener('click', () => select(tab.id));
			button.addEventListener('dblclick', () => setCollapsed(!root.hasAttribute('data-collapsed')));
			button.addEventListener('keydown', (event) => {
				const list = visibleTabs();
				const at = list.indexOf(button);
				const next =
					event.key === 'ArrowRight'
						? at + 1
						: event.key === 'ArrowLeft'
							? at - 1
							: event.key === 'Home'
								? 0
								: event.key === 'End'
									? list.length - 1
									: Number.NaN;
				if (Number.isNaN(next)) return;
				event.preventDefault();
				const target = list[(next + list.length) % list.length];
				target?.focus();
				if (target?.dataset.tab) select(target.dataset.tab);
			});
			tabsBar.append(button);
			const panel = el(doc, 'div', 'ribbon-panel');
			panel.id = `xve-panel-${tab.id}`;
			panel.dataset.tab = tab.id;
			panel.setAttribute('role', 'tabpanel');
			panel.setAttribute('aria-labelledby', button.id);
			for (const group of tab.groups) {
				const node = renderGroup(scope, group);
				if (node) panel.append(node);
			}
			root.append(panel);
		}
		tabsBar.append(collapse);
		refresh();
		if (!tabs.some((tab) => tab.id === selected)) selected = tabs[0]?.id ?? '';
		select(selected);
		stopOverflow = attachRibbonOverflow(root);
	};

	const refresh = () => {
		for (const update of scope.updates) update();
		for (const tab of tabs) {
			const button = tabButtons().find((item) => item.dataset.tab === tab.id);
			if (!button) continue;
			let show = true;
			try {
				show = tab.contextual ? tab.contextual(ctx) : true;
			} catch {
				show = false;
			}
			button.hidden = !show;
		}
		applyRibbonVisibility(root, handlers.isHidden);
		const current = tabButtons().find((tab) => tab.dataset.tab === selected);
		if (current && (current.hidden || current.hasAttribute('data-xve-hidden'))) {
			const fallback = visibleTabs()[0]?.dataset.tab;
			if (fallback) select(fallback);
		}
	};

	root.append(tabsBar);
	const stopListening = onRibbonTabsChange(() => build());
	build();
	return {
		element: root,
		refresh,
		rebuild: build,
		selectTab: select,
		activeTab: () => selected,
		showKeyTips: () => {
			const tab = tabButtons().find((item) => item.dataset.tab === selected) ?? fileTab;
			tab.focus();
			showTabKeyTips(root, tabsBar, fileTab, (id) => select(id));
		},
		focus: () => (tabButtons().find((item) => item.dataset.tab === selected) ?? fileTab).focus(),
		destroy: () => {
			stopListening();
			stopOverflow();
			closeRibbonPopover();
		},
	};
}
