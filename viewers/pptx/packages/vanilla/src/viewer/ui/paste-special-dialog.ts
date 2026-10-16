/**
 * Paste Special (Ctrl/Cmd+Alt+V): offers the four Paste Special formats
 * PowerPoint's own dialog offers (Keep Source Formatting, Use Destination
 * Theme, Picture, Keep Text Only), sourced from `pptx-viewer-shared` so the
 * option set and its labels cannot drift from the post-paste "Paste Options"
 * toolbar or the other four bindings.
 */
import type { PasteSpecialFormat } from 'ooxml-ui/pptx';
import { PASTE_SPECIAL_OPTIONS } from 'ooxml-ui/pptx';

import type { Translator } from '../i18n';
import { appendDialogFooter } from './dialog-footer';
import { appendRadioRow, createParityDialogShell } from './parity-dialog-shell';
import type { RegionSet } from './region-set';

export interface PasteSpecialDialogOptions {
	onConfirm(format: PasteSpecialFormat): void;
	/** The viewer's dialog regions: the dialog closes with the viewer. */
	regions?: RegionSet;
}

export function openPasteSpecialDialog(
	doc: Document,
	t: Translator,
	options: PasteSpecialDialogOptions,
): void {
	const shell = createParityDialogShell(
		doc,
		t,
		t('pptx.pasteSpecial.dialogTitle'),
		options.regions && { regions: options.regions },
	);
	let selected: PasteSpecialFormat = 'keep-source-formatting';
	appendRadioRow(
		doc,
		shell.body,
		'paste-special-format',
		PASTE_SPECIAL_OPTIONS.map((option) => ({ value: option.id, label: t(option.labelKey) })),
		selected,
		(value) => {
			selected = value as PasteSpecialFormat;
		},
	);
	appendDialogFooter(
		doc,
		shell.footer,
		[
			{ id: 'cancel', label: t('pptx.common.cancel') },
			{ id: 'ok', label: t('pptx.common.ok'), variant: 'primary' },
		],
		(id) => {
			shell.close();
			if (id === 'ok') {
				options.onConfirm(selected);
			}
		},
	);
}
