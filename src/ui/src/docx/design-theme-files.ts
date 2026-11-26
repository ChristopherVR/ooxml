import { loadThemePackage } from 'ooxml-core/drawingml';
import { designTheme, wordThemeXml } from 'ooxml-core/docx';
import { applyDesignEdit } from './design-commands';
import { dialogButton, labelled } from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import type { EditorCore } from './editor-core';
import { downloadBytes } from './file-commands';
import { translate } from './localization';
import { canEditDesign } from './design-ribbon-controls';

export function browseTheme(core: EditorCore, anchor: HTMLButtonElement): void {
	const root = anchor.getRootNode();
	if (!(root instanceof ShadowRoot)) return;
	const generation = core.loadGeneration;
	const input = document.createElement('input');
	input.type = 'file';
	input.accept = '.thmx';
	input.hidden = true;
	input.setAttribute('aria-label', translate(core.locale, 'design.browseThemes'));
	root.querySelector('input[data-theme-picker]')?.remove();
	input.dataset.themePicker = '';
	input.addEventListener('cancel', () => input.remove(), { once: true });
	input.addEventListener(
		'change',
		async () => {
			const file = input.files?.[0];
			input.remove();
			if (!file) return;
			try {
				const { MAX_THEME_PACKAGE_BYTES, readThemePackage } = await loadThemePackage();
				if (file.size > MAX_THEME_PACKAGE_BYTES) throw new Error();
				const theme = await readThemePackage(await file.arrayBuffer());
				if (generation === core.loadGeneration && anchor.isConnected)
					applyDesignEdit(core, { kind: 'importTheme', theme });
			} catch {
				if (generation === core.loadGeneration)
					core.host.reportError(translate(core.locale, 'design.importError'));
			}
		},
		{ once: true },
	);
	root.append(input);
	input.click();
}

export function saveTheme(core: EditorCore, anchor: HTMLButtonElement): void {
	const root = anchor.getRootNode();
	if (!(root instanceof ShadowRoot)) return;
	const theme = designTheme(core.model.theme);
	const generation = core.loadGeneration;
	const dialog = createDialogShell('design.saveTheme', 'dve-format-dialog dve-design-dialog');
	const input = document.createElement('input');
	input.type = 'text';
	input.maxLength = 100;
	input.value = (theme.name || translate(core.locale, 'Custom')).slice(0, 100);
	const cancel = dialogButton('Cancel');
	const save = dialogButton('Save', true);
	const close = () => {
		setDialogOpen(dialog, false);
		dialog.remove();
		anchor.focus();
	};
	const name = () =>
		input.value
			.trim()
			.replace(/\p{Cc}|[<>:"/\\|?*]/gu, '')
			.replace(/\.thmx$/iu, '')
			.replace(/[. ]+$/u, '');
	input.addEventListener('input', () => {
		save.disabled = !name();
	});
	save.disabled = !name();
	cancel.addEventListener('click', close);
	save.addEventListener('click', async () => {
		const themeName = name();
		if (!themeName || !canEditDesign(core) || generation !== core.loadGeneration) return;
		close();
		try {
			const { writeThemePackage } = await loadThemePackage();
			downloadBytes(
				await writeThemePackage(wordThemeXml({ ...theme, name: themeName })),
				`${themeName}.thmx`,
			);
		} catch {
			core.host.reportError(translate(core.locale, 'design.exportError'));
		}
	});
	onDialogDismiss(dialog, close);
	dialog.append(labelled('design.schemeName', input), dialogActions(cancel, save));
	localizeDialog(dialog, core.locale);
	root.append(dialog);
	setDialogOpen(dialog, true);
}
