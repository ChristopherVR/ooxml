import type { VisioStencilEdit } from 'ooxml-core/visio';
import { createStencil, editStencil } from './stencil-client';
import { ViewerDialog } from './viewer-dialog';
import { MY_SHAPES_ID, MY_SHAPES_NAME, type StencilFiles } from './stencil-files';

type Translate = (text: string) => string;

export type StencilAction =
	| { type: 'new' }
	| { type: 'collect'; draft: () => VisioStencilEdit }
	| { type: 'save' | 'add-selection'; stencil: string; draft?: (name: string) => VisioStencilEdit }
	| { type: 'rename' | 'edit' | 'delete'; stencil: string; master: string };

export async function runStencilAction(
	pane: HTMLElement,
	files: StencilFiles,
	action: StencilAction,
	t: Translate,
	update: (id: string) => void,
): Promise<void> {
	if (action.type === 'new') {
		const name = await ask(pane, 'New Stencil', 'Stencil name', t('Untitled Stencil'), t);
		if (name === undefined) return;
		const stencil = await files.open(`${name.trim()}.vssx`, await createStencil());
		update(stencil.id);
		return;
	}
	if (action.type === 'collect') {
		const edit = action.draft();
		const existing = files.get(MY_SHAPES_ID);
		const source = existing?.bytes ?? (await createStencil());
		const result = await editStencil(source, edit);
		await files.replace(MY_SHAPES_ID, MY_SHAPES_NAME, result.bytes);
		update(MY_SHAPES_ID);
		return;
	}
	const stencil = files.get(action.stencil);
	if (!stencil?.editable) return;
	if (action.type === 'save') {
		const { bytes, extension } = await editStencil(stencil.bytes);
		const name = stencil.id === MY_SHAPES_ID ? t(stencil.name) : stencil.name;
		download(pane, bytes, `${name.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')}.${extension}`);
		return;
	}
	let edit: VisioStencilEdit | undefined;
	if (action.type === 'add-selection') {
		const draft = action.draft?.('');
		if (!draft) return;
		if (draft.type === 'add-shape') {
			const name = await ask(pane, 'Add Selected Shape', 'Master name', t('New Master'), t);
			if (name !== undefined) edit = { ...draft, name };
		} else edit = draft;
	} else {
		const file = 'master' in action ? files.fileMaster(action.master) : undefined;
		if (!file) return;
		const { master } = file;
		if (action.type === 'rename') {
			const name = await ask(pane, 'Rename Master', 'Master name', master.name, t);
			if (name !== undefined) edit = { type: 'rename-master', id: master.id, name };
		} else if (action.type === 'delete') {
			edit = { type: 'delete-master', id: master.id };
		} else {
			const shape = master.shapes[0];
			if (!shape) return;
			const text = await ask(pane, 'Edit Master', 'Default shape text', shape.text.plainText, t);
			if (text !== undefined)
				edit = { type: 'edit-master-text', id: master.id, shapeId: shape.id, text };
		}
	}
	if (!edit || files.get(stencil.id) !== stencil) return;
	const result = await editStencil(stencil.bytes, edit);
	if (files.get(stencil.id) !== stencil) return;
	await files.replace(stencil.id, stencil.name, result.bytes);
	update(stencil.id);
}

function ask(
	pane: HTMLElement,
	heading: string,
	label: string,
	initial: string,
	t: Translate,
): Promise<string | undefined> {
	return new Promise((resolve) => {
		const dialog = new ViewerDialog(
			pane,
			'stencil-dialog',
			t(heading),
			[t('OK'), t('Cancel')],
			(button) => {
				if (button === t('OK') && !isText && !input.value.trim()) {
					dialog.error.textContent = t('Enter a name.');
					return;
				}
				resolve(button === t('OK') ? input.value : undefined);
				dialog.close();
			},
		);
		const field = pane.ownerDocument.createElement('label');
		const isText = heading === 'Edit Master';
		const input = pane.ownerDocument.createElement(isText ? 'textarea' : 'input');
		input.maxLength = isText ? 4096 : 255;
		input.value = initial;
		input.setAttribute('aria-label', t(label));
		input.addEventListener('keydown', (event) => {
			if (!isText && 'key' in event && event.key === 'Enter') {
				event.preventDefault();
				if (!input.value.trim()) {
					dialog.error.textContent = t('Enter a name.');
					return;
				}
				resolve(input.value);
				dialog.close();
			}
		});
		field.textContent = t(label);
		field.append(input);
		dialog.body.append(field);
		dialog.dialog.setAttribute('close-label', t('Close'));
		dialog.dialog.addEventListener(
			'office-dialog-close',
			() => {
				resolve(undefined);
				queueMicrotask(() => dialog.dialog.remove());
			},
			{ once: true },
		);
		dialog.show();
	});
}

function download(pane: HTMLElement, bytes: Uint8Array, name: string): void {
	const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'application/zip' }));
	const anchor = pane.ownerDocument.createElement('a');
	anchor.href = url;
	anchor.download = name;
	anchor.hidden = true;
	pane.append(anchor);
	anchor.click();
	anchor.remove();
	setTimeout(() => URL.revokeObjectURL(url), 1000);
}
