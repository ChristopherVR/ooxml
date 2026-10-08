/**
 * Excel-style title bar: the X badge and quick access toolbar (Save, Undo, Redo), the file name
 * and save state and "Tell me". The shared `office-ui-title-bar` draws them from translated
 * state; while the workbook is shared, the people in the session (the shared presence stack) sit
 * in its `collaboration` slot. The editing mode, Comments and Share live at the right end of the
 * ribbon tab row, as in Excel (ribbon-actions.ts).
 */
import { defineTitleBar } from '../controls';
import type { OfficeTitleBarState } from '../controls';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { historyOf, searchCommands, type TellMeHandlers } from 'ooxml-core/xlsx/ui';
import { definePresence, type PresenceParticipant } from '../presence';
import { participants } from './backstage/pages-share';
import type { XlsxCollaborationState } from './collaboration-types';

export type SaveState = 'saved' | 'dirty' | 'saving' | 'saved-local';

export interface TitleBarHandlers extends TellMeHandlers {
	save(): void;
	/** Sharing state for the presence stack; absent hides it. */
	collaboration?(): XlsxCollaborationState;
}

export interface TitleBar {
	readonly element: HTMLElement;
	setFileName(name: string): void;
	setSaveState(state: SaveState): void;
	/** Re-reads undo/redo and the people in a shared session. */
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

	definePresence();
	const people = doc.createElement('office-ui-presence') as HTMLElement & {
		participants: PresenceParticipant[];
	};
	people.className = 'xve-title-people';
	people.slot = 'collaboration';
	people.setAttribute('max', '3');
	people.hidden = true;
	element.append(people);

	let fileName = '';
	let saveState: SaveState = 'saved';

	const render = () => {
		const undoLabel = historyOf(ctx)?.undoLabel();
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

	const relocalize = () => {
		people.setAttribute('label', ctx.t('People in this session'));
		render();
	};
	const refresh = () => {
		const sharing = handlers.collaboration?.();
		people.hidden = !sharing?.active;
		const next = sharing?.active ? participants(sharing) : [];
		if (JSON.stringify(next) !== JSON.stringify(people.participants)) people.participants = next;
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
