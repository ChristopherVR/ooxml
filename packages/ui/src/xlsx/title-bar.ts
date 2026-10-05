/**
 * Excel-style title bar: the X badge and quick access toolbar (Save, Undo, Redo), the file name
 * and save state, "Tell me", the comments toggle and the Editing / Viewing select. The shared
 * `office-ui-title-bar` draws the mark, Quick Access Toolbar, name and command search from
 * translated state; the comments toggle and mode select are product controls in its `actions` slot.
 */
import { defineTitleBar } from '../controls';
import type { OfficeTitleBarState } from '../controls';
import type { EditorContext } from './context';
import { el } from './ribbon/controls';
import { ribbonIcon } from './ribbon/icons';
import { searchCommands, type TellMeHandlers } from './ribbon/tell-me';

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

type TitleBarElement = HTMLElement & {
	state: OfficeTitleBarState;
	searchField: HTMLElement | null;
};

export function createTitleBar(ctx: EditorContext, handlers: TitleBarHandlers): TitleBar {
	defineTitleBar();
	const doc = ctx.host.ownerDocument;
	const element = doc.createElement('office-ui-title-bar') as TitleBarElement;
	element.className = 'xve-titlebar';
	element.setAttribute('part', 'title-bar');

	const actions = el(doc, 'div', 'xve-title-actions');
	actions.slot = 'actions';
	const comments = el(doc, 'button', 'xve-icon-button');
	comments.type = 'button';
	comments.append(ribbonIcon(doc, 'comments', 16));
	comments.addEventListener('click', () => void ctx.commands.run('review.show-comments'));
	const mode = el(doc, 'select', 'xve-mode-select');
	mode.append(new Option('', 'editing'), new Option('', 'viewing'));
	mode.addEventListener('change', () => handlers.setReadOnly(mode.value === 'viewing'));
	actions.append(comments, mode);
	element.append(actions);

	let fileName = '';
	let saveState: SaveState = 'saved';

	const render = () => {
		const session = ctx.session();
		const undoLabel = session?.undoLabel();
		const entries = [
			{ id: 'save', icon: 'save', label: ctx.t('Save'), title: `${ctx.t('Save')} (Ctrl+S)` },
			{
				id: 'undo',
				icon: 'undo',
				label: ctx.t('Undo'),
				title: `${ctx.t('Undo')}${undoLabel ? ` ${ctx.t(undoLabel)}` : ''} (Ctrl+Z)`,
				disabled: !ctx.commands.isEnabled('edit.undo'),
			},
			{
				id: 'redo',
				icon: 'redo',
				label: ctx.t('Redo'),
				title: `${ctx.t('Redo')} (Ctrl+Y)`,
				disabled: !ctx.commands.isEnabled('edit.redo'),
			},
		];
		element.state = {
			appMark: 'X',
			fileName,
			status: ctx.t(SAVE_STATE_TEXT[saveState]),
			tone: saveState === 'saving' ? 'saving' : 'idle',
			quickAccess: { label: ctx.t('Quick access'), items: entries },
			search: {
				placeholder: ctx.t('Tell me what you want to do'),
				label: ctx.t('Tell me what you want to do'),
				heading: ctx.t('Best matches'),
				empty: ctx.t('No matching commands'),
				match: (query) =>
					searchCommands(ctx, query, handlers.isHidden).map((command) => ({
						id: command.id,
						label: ctx.t(command.label),
						...(command.shortcut ? { category: command.shortcut } : {}),
					})),
			},
		};
	};

	element.addEventListener('office-command', (event) => {
		const id = (event as CustomEvent<{ command: string }>).detail.command;
		if (id === 'save') handlers.save();
		else if (id === 'undo') void ctx.commands.run('edit.undo');
		else if (id === 'redo') void ctx.commands.run('edit.redo');
	});
	element.addEventListener('office-command-search', (event) => {
		const { command: id } = (event as CustomEvent<{ query: string; command?: string }>).detail;
		const command = id ? ctx.commands.get(id) : undefined;
		if (!command) return;
		// A command that needs a value takes the user to its ribbon control.
		if (command.value && handlers.revealControl(command.id)) return;
		ctx.grid()?.focus();
		void ctx.commands.run(command.id);
	});

	const label = (control: HTMLElement, text: string) => {
		const translated = ctx.t(text);
		control.setAttribute('aria-label', translated);
		control.title = translated;
	};
	const relocalize = () => {
		label(comments, 'Show comments');
		mode.setAttribute('aria-label', ctx.t('Editing mode'));
		mode.options[0]!.textContent = ctx.t('Editing');
		mode.options[1]!.textContent = ctx.t('Viewing');
		render();
	};
	const refresh = () => {
		mode.value = ctx.readOnly() ? 'viewing' : 'editing';
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
		render();
	};
	relocalize();
	return {
		element,
		setFileName(value) {
			fileName = value;
			render();
		},
		setSaveState(value) {
			saveState = value;
			render();
		},
		refresh,
		relocalize,
		focusTellMe: () => element.searchField?.focus(),
	};
}
