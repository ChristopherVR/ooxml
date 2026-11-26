/**
 * Builds the shadow-DOM shell once: title bar, ribbon, formula bar row, grid, sheet tabs, status
 * bar, backstage, shortcut help and toasts; installs the commands and mounts the grid modules.
 * Mirrors docx-viewer's editor-shell.ts.
 */
import {
	createBackstage,
	type Backstage,
	type BackstageHost,
	type IterationSettings,
} from './backstage';
import type { EditorCore } from './editor-core';
import {
	openCloudFile,
	reportError,
	runFileCommand,
	templateInto,
	type FileChrome,
} from './editor-files';
import { emit, type FileCommand } from './events';
import { installKeyboard, type ShellAction } from './keyboard';
import {
	installCommands,
	mountFormulaBar,
	mountGrid,
	mountPageViews,
	mountSheetTabs,
} from './modules';
import { createRibbon, type Ribbon } from './ribbon/ribbon';
import { createCommentsPaneSlot } from './comments-pane-slot';
import { createRibbonActions } from './ribbon-actions';
import { el } from './ribbon/controls';
import { hideKeyTips } from './ribbon/keytips';
import { shellCommands } from './shell-commands';
import { createShortcutHelp } from './shortcut-help';
import { commandsOnlyInTabs, ribbonTabs } from 'ooxml-core/xlsx/ui';
import { createStatusBar } from './status-bar';
import { editorStyleText } from './theme';
import { createTitleBar, type SaveState, type TitleBar } from './title-bar';
import { createToaster } from './toast';
import { packagePassword } from './dialogs/package-password';
import { attachShareDialog, type ShareDialogController } from '../dialog/share-controller';
import { modalFocusableElements } from '../dialog/focus';
import { createSelectionPane } from './selection-pane';
import { registerChartPanes } from './chart-panes';
import { createSmartArtController } from './smartart-controller';

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
	/** The Share dialog (Share button, File > Share, `element.share()`). */
	readonly share: ShareDialogController;
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
	const isHidden = (id: string) => core.isHidden(id);
	const chrome: FileChrome = {
		setSaveState(state, detail) {
			saveState = state;
			titleBar.setSaveState(state, detail);
		},
		fileNameChanged: () => titleBar.setFileName(core.fileName),
	};
	const fileCommand = (command: FileCommand, fileName?: string) =>
		runFileCommand(core, command, fileName, {
			chrome,
			openBackstage: (page) => backstage.open(page),
			// Over File > Share when it is open, as in Office; its summary follows the session.
			openShare: () => share.open(),
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
		fileProviders: () => core.fileProviders,
		openCloudFile(file) {
			void openCloudFile(core, file, chrome, closeBackstage).catch((cause) =>
				reportError(core, cause),
			);
		},
		close: closeBackstage,
		options: () => ({
			locale: core.locale,
			theme: core.theme,
			author: core.authorName,
			calculation: core.calculation,
			iteration: core.iteration,
			allowFontDownloads: core.webfonts.allowed(),
			fontDownloadsLocked: core.webfonts.lock() !== undefined,
		}),
		setOption(key, raw) {
			if (key === 'iteration') return core.setIterativeCalculation(raw as IterationSettings);
			if (key === 'allowFontDownloads') {
				if (core.webfonts.lock() === undefined) core.webfonts.setReaderAllows(raw === true);
				return;
			}
			if (key === 'fontDownloadsLocked') return;
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
		share: {
			state: () => core.collab.state(),
			open: () => void fileCommand('share'),
			stop: () => core.collab.stop(),
		},
		hiddenActions: () => core.hiddenActions,
		customization: () => core.customization,
		regionContext: (region) => ({
			region,
			editor: element,
			fileName: core.fileName,
			locale: core.locale,
			canEdit: !core.readOnly,
			activeSheet: core.activeSheet,
			close: closeBackstage,
		}),
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
	const ribbon = createRibbon(ctx, {
		openBackstage: () => backstage.open('info'),
		isHidden,
		hiddenTabs: () => core.customization.hiddenTabs,
	});
	// Editing mode, Comments and Share at the right end of the tab row, as in Excel.
	const ribbonActions = createRibbonActions(ctx, {
		isHidden,
		setReadOnly: (readOnly) => (element.readOnly = readOnly),
		collaboration: () => core.collab.state(),
	});
	ribbon.element.append(ribbonActions.element);
	const titleBar = createTitleBar(ctx, {
		isHidden,
		searchHidden: () => commandsOnlyInTabs(ribbonTabs(), core.customization.hiddenTabs),
		save: () => void fileCommand('save'),
		collaboration: () => core.collab.state(),
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
		infoAvailable: () => !core.customization.fileHidden,
		setZoom: (percent) => {
			if (ctx.commands.get('view.set-zoom')) void ctx.commands.run('view.set-zoom', percent);
			else ctx.grid()?.setZoom(percent);
			core.requestRender();
		},
	});
	const toaster = createToaster(doc, () => ctx.t('Close'));

	const workspace = el(doc, 'div', 'xve-workspace');
	const selectionPane = createSelectionPane(ctx);
	const commentsPane = createCommentsPaneSlot(ctx);
	const smartArt = createSmartArtController(ctx);
	const canvasRow = el(doc, 'div', 'xve-canvas-row');
	const formulaRow = el(doc, 'div', 'xve-formula-row');
	const gridHost = el(doc, 'div', 'xve-grid-host');
	const tabsRow = el(doc, 'div', 'xve-tabs-row');
	const empty = el(doc, 'div', 'xve-empty');
	const emptyText = el(doc, 'p');
	empty.append(emptyText);
	workspace.append(formulaRow, gridHost, tabsRow, empty);
	canvasRow.append(workspace, smartArt.element, selectionPane.element, commentsPane.element);
	const chartPanes = registerChartPanes(ctx, workspace, smartArt.element);
	frame.append(
		titleBar.element,
		ribbon.element,
		statusBar.signatureNotice,
		canvasRow,
		statusBar.element,
		backstage.element,
		backstage.optionsDialog,
		help.element,
		toaster.element,
	);
	root.append(style, frame);
	titleBar.setFileName(core.fileName);

	safeMount('formula bar', () => mountFormulaBar(ctx, formulaRow));
	safeMount('grid', () => mountGrid(ctx, gridHost));
	let pageViews: { refresh(): void } | undefined;
	safeMount('page views', () => {
		pageViews = mountPageViews(ctx, gridHost, () => core.fileName);
	});
	safeMount('sheet tabs', () => mountSheetTabs(ctx, tabsRow));
	// Parts the browser specs and hosts style; set on the containers when a module did not.
	if (!gridHost.querySelector('[part~="grid"]')) gridHost.setAttribute('part', 'grid');
	if (!tabsRow.querySelector('[part~="sheet-tabs"]')) tabsRow.setAttribute('part', 'sheet-tabs');
	if (!formulaRow.querySelector('[part~="formula-bar"]'))
		formulaRow.setAttribute('part', 'formula-bar');

	const regions = () =>
		[
			ribbon.element,
			formulaRow,
			gridHost,
			tabsRow,
			...chartPanes.elements(),
			smartArt.element,
			selectionPane.element,
			commentsPane.element,
			statusBar.element,
		].filter((region) => region.getClientRects().length);
	const focusRegion = (region: HTMLElement) => {
		if (region === ribbon.element) return ribbon.focus();
		if (region === gridHost) return ctx.grid()?.focus();
		// The first stop that is laid out: a hidden button (no compatibility notes) takes no focus.
		modalFocusableElements(region)
			.find((stop) => stop.getClientRects().length)
			?.focus();
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
		const from = at >= 0 ? at : step > 0 ? -1 : list.length;
		hideKeyTips();
		// A region with nothing to focus is passed over, so F6 never stops on one.
		for (let i = 1; i <= list.length; i++) {
			const region = list[(from + step * i + list.length * 2) % list.length]!;
			focusRegion(region);
			const now = root.activeElement;
			if (now && region.contains(now)) return;
		}
	};
	installKeyboard(ctx, actions);
	const share = attachShareDialog({
		container: root,
		host: {
			state: () => core.collab.state(),
			start: (options) => core.collab.start(options),
			stop: () => core.collab.stop(),
		},
		roomPrefix: 'book',
		defaults: () => ({ userName: core.authorName }),
		locale: () => core.locale,
		note: () =>
			ctx.t(
				'Cell values, formulas, formats, sheets, rows, columns, merges and names are shared. Charts, pictures, comments, conditional formats, data validation, tables and page setup stay in this window.',
			),
	});
	element.addEventListener('collaboration-change', () => {
		backstage.refresh('share');
		share.refresh();
	});

	core.shell = {
		render() {
			ribbon.refresh();
			titleBar.refresh();
			ribbonActions.refresh();
			statusBar.refresh();
			pageViews?.refresh();
			chartPanes.refresh();
			smartArt.refresh();
			selectionPane.refresh();
			commentsPane.refresh();
			titleBar.setFileName(core.fileName);
			if (core.dirty.dirty && saveState !== 'dirty' && saveState !== 'saving')
				chrome.setSaveState('dirty');
			if (!core.dirty.dirty && saveState === 'dirty') chrome.setSaveState('saved');
			empty.hidden = Boolean(core.workbook);
			// Without a File tab (`ribbon.hiddenTabs`) the host opens workbooks: do not point at File.
			const message = ctx.t(
				core.customization.fileHidden
					? 'No workbook is open.'
					: 'Open a workbook or create a new one from the File tab.',
			);
			if (emptyText.textContent !== message) emptyText.textContent = message;
		},
		toast: (message, kind) => toaster.show(message, kind),
		relocalize() {
			element.setAttribute('aria-label', ctx.t('Spreadsheet editor'));
			ribbon.rebuild();
			titleBar.relocalize();
			ribbonActions.relocalize();
			statusBar.relocalize();
			backstage.relocalize();
			share.relocalize();
			chartPanes.relocalize();
			smartArt.relocalize();
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
		share,
	};
}
