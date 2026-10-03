/**
 * Builds the shadow-DOM shell once: title bar, ribbon, formula bar row, grid, sheet tabs, status
 * bar, backstage, shortcut help and toasts; installs the commands and mounts the grid modules.
 * Mirrors docx-viewer's editor-shell.ts.
 */
import { createBackstage, type Backstage, type BackstageHost } from './backstage';
import type { EditorCore } from './editor-core';
import { runFileCommand, templateInto, type FileChrome } from './editor-files';
import { emit, type FileCommand } from './events';
import { installKeyboard, type ShellAction } from './keyboard';
import { installCommands, mountFormulaBar, mountGrid, mountSheetTabs } from './modules';
import { createRibbon, type Ribbon } from './ribbon/ribbon';
import { el } from './ribbon/controls';
import { hideKeyTips } from './ribbon/keytips';
import { shellCommands } from './shell-commands';
import { createShortcutHelp } from './shortcut-help';
import { createStatusBar } from './status-bar';
import { editorStyleText } from './theme';
import { createTitleBar, type SaveState, type TitleBar } from './title-bar';
import { createToaster } from './toast';
import { packagePassword } from './dialogs/package-password';

/** Element properties the shell's options and customize pages change. */
export interface ShellElement extends HTMLElement {
	locale: string;
	theme: 'light' | 'dark' | 'auto';
	authorName: string;
	readOnly: boolean;
	showFormulaBar: boolean;
	hiddenActions: string[];
}

export interface Shell {
	readonly ribbon: Ribbon;
	readonly titleBar: TitleBar;
	readonly backstage: Backstage;
	readonly chrome: FileChrome;
	saveState(): SaveState;
	fileCommand(command: FileCommand, fileName?: string): Promise<void>;
}

function safeMount(name: string, mount: () => unknown): void {
	try {
		mount();
	} catch (error) {
		console.error(`xlsx-editor: ${name} failed to mount`, error);
	}
}

export function buildShell(core: EditorCore): Shell {
	const element = core.host as ShellElement;
	const { ctx } = core;
	const doc = element.ownerDocument;
	const root = element.shadowRoot ?? element.attachShadow({ mode: 'open' });
	const style = doc.createElement('style');
	style.textContent = editorStyleText;
	const frame = el(doc, 'section', 'xve-frame');
	let saveState: SaveState = 'saved';
	const isHidden = (id: string) => core.hiddenActions.includes(id);
	const chrome: FileChrome = {
		setSaveState(state) {
			saveState = state;
			titleBar.setSaveState(state);
		},
		fileNameChanged: () => titleBar.setFileName(core.fileName),
	};
	const fileCommand = (command: FileCommand, fileName?: string) =>
		runFileCommand(core, command, fileName, {
			chrome,
			openBackstage: (page) => backstage.open(page),
		});
	const closeBackstage = () => {
		backstage.close();
		ribbon.focus();
	};
	const backstageHost: BackstageHost = {
		ctx,
		fileName: () => core.fileName,
		saveState: () => saveState,
		fileCommand: (command, fileName) => void fileCommand(command, fileName),
		newFromTemplate: (id) => templateInto(core, id, chrome),
		close: closeBackstage,
		options: () => ({
			locale: core.locale,
			theme: core.theme,
			author: core.authorName,
			calculation: core.calculation,
		}),
		setOption(key, raw) {
			const value = String(raw);
			if (key === 'locale') element.locale = value;
			else if (key === 'theme')
				element.theme = value === 'light' || value === 'dark' ? value : 'auto';
			else if (key === 'author') element.authorName = value;
			else core.setCalculation(value === 'manual' ? 'manual' : 'automatic');
		},
		setProperty(key, value) {
			if (core.readOnly) return;
			core.session?.setDocumentProperties({ [key]: value === '' ? null : value });
		},
		passwordProtected: () => core.savePassword !== undefined,
		async setPassword() {
			if (core.readOnly || !core.workbook) return;
			const workbook = core.workbook;
			const password = await packagePassword(ctx, true);
			if (password === undefined || core.workbook !== workbook || core.readOnly) return;
			core.savePassword = password;
			core.dirty.set(true);
			backstage.relocalize();
		},
		removePassword() {
			if (core.readOnly) return;
			core.savePassword = undefined;
			core.dirty.set(true);
			backstage.relocalize();
		},
		hiddenActions: () => core.hiddenActions,
		setHiddenActions(ids) {
			element.hiddenActions = ids;
			emit(element, 'ribbon-customize', { hiddenActions: [...ids] });
		},
	};
	const backstage = createBackstage(backstageHost);
	const help = createShortcutHelp(ctx, () => ctx.grid()?.focus());
	ctx.commands.registerAll(
		shellCommands({
			fileCommand: (command) => void fileCommand(command),
			openBackstage: (page) => backstage.open(page),
			showShortcuts: () => help.open(),
			formulaBarShown: () => element.showFormulaBar,
			setFormulaBarShown: (shown) => (element.showFormulaBar = shown),
		}),
	);
	safeMount('commands', () => installCommands(ctx));
	const ribbon = createRibbon(ctx, { openBackstage: () => backstage.open('info'), isHidden });
	const titleBar = createTitleBar(ctx, {
		isHidden,
		save: () => void fileCommand('save'),
		setReadOnly: (readOnly) => (element.readOnly = readOnly),
		revealControl(id) {
			const control = ribbon.element.querySelector<HTMLElement>(
				`.ribbon-panel [data-command="${id.replace(/["\\]/g, '')}"]`,
			);
			const tab = control?.closest<HTMLElement>('.ribbon-panel')?.dataset.tab;
			if (!control || !tab) return false;
			ribbon.selectTab(tab);
			(control.querySelector('input') ?? control).focus();
			return true;
		},
	});
	const statusBar = createStatusBar(ctx, {
		showNotes: () => backstage.open('info'),
		setZoom: (percent) => {
			if (ctx.commands.get('view.set-zoom')) void ctx.commands.run('view.set-zoom', percent);
			else ctx.grid()?.setZoom(percent);
			core.requestRender();
		},
	});
	const toaster = createToaster(doc, () => ctx.t('Close'));

	const workspace = el(doc, 'div', 'xve-workspace');
	const formulaRow = el(doc, 'div', 'xve-formula-row');
	const gridHost = el(doc, 'div', 'xve-grid-host');
	const tabsRow = el(doc, 'div', 'xve-tabs-row');
	const empty = el(doc, 'div', 'xve-empty');
	const emptyText = el(doc, 'p');
	empty.append(emptyText);
	workspace.append(formulaRow, gridHost, tabsRow, empty);
	frame.append(
		titleBar.element,
		ribbon.element,
		workspace,
		statusBar.element,
		backstage.element,
		help.element,
		toaster.element,
	);
	root.append(style, frame);
	titleBar.setFileName(core.fileName);

	safeMount('formula bar', () => mountFormulaBar(ctx, formulaRow));
	safeMount('grid', () => mountGrid(ctx, gridHost));
	safeMount('sheet tabs', () => mountSheetTabs(ctx, tabsRow));
	// Parts the browser specs and hosts style; set on the containers when a module did not.
	if (!gridHost.querySelector('[part~="grid"]')) gridHost.setAttribute('part', 'grid');
	if (!tabsRow.querySelector('[part~="sheet-tabs"]')) tabsRow.setAttribute('part', 'sheet-tabs');
	if (!formulaRow.querySelector('[part~="formula-bar"]'))
		formulaRow.setAttribute('part', 'formula-bar');

	const regions = () =>
		[ribbon.element, formulaRow, gridHost, tabsRow, statusBar.element].filter(
			(region) => region.getClientRects().length,
		);
	const focusRegion = (region: HTMLElement) => {
		if (region === ribbon.element) return ribbon.focus();
		if (region === gridHost) return ctx.grid()?.focus();
		region.querySelector<HTMLElement>('input, button:not(:disabled), [tabindex="0"]')?.focus();
	};
	const actions = (action: ShellAction) => {
		if (action === 'help') return help.open();
		if (action === 'keytips') return ribbon.showKeyTips();
		if (action === 'context-menu') {
			const target = (root.activeElement as HTMLElement | null) ?? gridHost;
			const box = target.getBoundingClientRect();
			target.dispatchEvent(
				new MouseEvent('contextmenu', {
					bubbles: true,
					composed: true,
					cancelable: true,
					clientX: box.left + 8,
					clientY: box.top + 8,
				}),
			);
			return;
		}
		const list = regions();
		const active = root.activeElement;
		const at = list.findIndex((region) => active && region.contains(active));
		const step = action === 'next-region' ? 1 : -1;
		const next = list[(at + step + list.length) % list.length] ?? gridHost;
		hideKeyTips();
		focusRegion(next);
	};
	installKeyboard(ctx, actions);

	core.shell = {
		render() {
			ribbon.refresh();
			titleBar.refresh();
			statusBar.refresh();
			titleBar.setFileName(core.fileName);
			if (core.dirty.dirty && saveState !== 'dirty' && saveState !== 'saving')
				chrome.setSaveState('dirty');
			if (!core.dirty.dirty && saveState === 'dirty') chrome.setSaveState('saved');
			empty.hidden = Boolean(core.workbook);
		},
		toast: (message, kind) => toaster.show(message, kind),
		relocalize() {
			element.setAttribute('aria-label', ctx.t('Spreadsheet editor'));
			emptyText.textContent = ctx.t('Open a workbook or create a new one from the File tab.');
			ribbon.rebuild();
			titleBar.relocalize();
			statusBar.relocalize();
			backstage.relocalize();
			core.requestRender();
		},
	};
	element.setAttribute('role', 'region');
	core.shell.relocalize();
	return {
		ribbon,
		titleBar,
		backstage,
		chrome,
		saveState: () => saveState,
		fileCommand,
	};
}
