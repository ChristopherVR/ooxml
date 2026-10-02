/**
 * Excel-style title bar: the X badge and quick access toolbar (Save, Undo, Redo), the file name
 * and save state, "Tell me", the comments toggle and the Editing / Viewing select. Layout mirrors
 * docx-viewer's title-bar.ts.
 */
import type { EditorContext } from './context';
import { el } from './ribbon/controls';
import { ribbonIcon } from './ribbon/icons';
import { createTellMe, type TellMeHandlers } from './ribbon/tell-me';

export type SaveState = 'saved' | 'dirty' | 'saving' | 'saved-local';

export interface TitleBarHandlers extends TellMeHandlers {
	save(): void;
	setReadOnly(readOnly: boolean): void;
}

export interface TitleBar {
	readonly element: HTMLElement;
	setFileName(name: string): void;
	setSaveState(state: SaveState): void;
	/** Re-reads read-only, undo/redo and comments state. */
	refresh(): void;
	relocalize(): void;
	focusTellMe(): void;
}

const SAVE_STATE_TEXT: Record<SaveState, string> = {
	saved: 'Saved',
	dirty: 'Unsaved changes',
	saving: 'Saving...',
	'saved-local': 'Saved to this device',
};

export function createTitleBar(ctx: EditorContext, handlers: TitleBarHandlers): TitleBar {
	const doc = ctx.host.ownerDocument;
	const element = el(doc, 'header', 'xve-titlebar');
	element.setAttribute('part', 'title-bar');
	const quick = el(doc, 'div', 'xve-quick-access');
	quick.setAttribute('role', 'toolbar');
	const badge = el(doc, 'span', 'xve-app-badge');
	badge.textContent = 'X';
	badge.setAttribute('aria-hidden', 'true');
	const iconButton = (icon: string) => {
		const button = el(doc, 'button', 'xve-icon-button');
		button.type = 'button';
		button.append(ribbonIcon(doc, icon, 16));
		button.addEventListener('mousedown', (event) => event.preventDefault());
		return button;
	};
	const save = iconButton('save');
	save.addEventListener('click', () => handlers.save());
	const undo = iconButton('undo');
	undo.addEventListener('click', () => void ctx.commands.run('edit.undo'));
	const redo = iconButton('redo');
	redo.addEventListener('click', () => void ctx.commands.run('edit.redo'));
	quick.append(badge, save, undo, redo);

	const title = el(doc, 'div', 'xve-document-title');
	const name = el(doc, 'span', 'xve-filename');
	const separator = el(doc, 'span', 'xve-title-separator');
	separator.textContent = '-';
	separator.setAttribute('aria-hidden', 'true');
	const state = el(doc, 'span', 'xve-save-state');
	state.setAttribute('aria-live', 'polite');
	title.append(name, separator, state);
	const start = el(doc, 'div', 'xve-titlebar-start');
	start.append(quick, title);

	const tellMe = createTellMe(ctx, handlers);
	const actions = el(doc, 'div', 'xve-title-actions');
	const comments = iconButton('comments');
	comments.addEventListener('click', () => void ctx.commands.run('review.show-comments'));
	const mode = el(doc, 'select', 'xve-mode-select');
	mode.append(new Option('', 'editing'), new Option('', 'viewing'));
	mode.addEventListener('change', () => handlers.setReadOnly(mode.value === 'viewing'));
	actions.append(comments, mode);
	element.append(start, tellMe.element, actions);

	let saveState: SaveState = 'saved';
	const label = (button: HTMLElement, text: string, shortcut?: string) => {
		const translated = ctx.t(text);
		button.setAttribute('aria-label', translated);
		button.title = shortcut ? `${translated} (${shortcut})` : translated;
	};
	const relocalize = () => {
		quick.setAttribute('aria-label', ctx.t('Quick access'));
		label(save, 'Save', 'Ctrl+S');
		label(undo, 'Undo', 'Ctrl+Z');
		label(redo, 'Redo', 'Ctrl+Y');
		label(comments, 'Show comments');
		mode.setAttribute('aria-label', ctx.t('Editing mode'));
		mode.options[0]!.textContent = ctx.t('Editing');
		mode.options[1]!.textContent = ctx.t('Viewing');
		state.textContent = ctx.t(SAVE_STATE_TEXT[saveState]);
		tellMe.relocalize();
	};
	const refresh = () => {
		mode.value = ctx.readOnly() ? 'viewing' : 'editing';
		undo.disabled = !ctx.commands.isEnabled('edit.undo');
		redo.disabled = !ctx.commands.isEnabled('edit.redo');
		const session = ctx.session();
		const undoLabel = session?.undoLabel();
		undo.title = `${ctx.t('Undo')}${undoLabel ? ` ${ctx.t(undoLabel)}` : ''} (Ctrl+Z)`;
		const showComments = ctx.commands.get('review.show-comments');
		comments.hidden = !showComments;
		if (showComments) {
			let pressed = false;
			try {
				pressed = showComments.checked?.(ctx) ?? false;
			} catch {
				pressed = false;
			}
			comments.setAttribute('aria-pressed', String(pressed));
		}
	};
	relocalize();
	return {
		element,
		setFileName(value) {
			name.textContent = value;
			name.title = value;
		},
		setSaveState(value) {
			saveState = value;
			state.dataset.state = value;
			state.textContent = ctx.t(SAVE_STATE_TEXT[value]);
		},
		refresh,
		relocalize,
		focusTellMe: () => tellMe.focus(),
	};
}
