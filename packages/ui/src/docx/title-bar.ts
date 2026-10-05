import { defineSearchField, defineSwitch, defineTitleBar } from '../controls';
import type { OfficeTitleBarCommand, OfficeTitleBarState } from '../controls';
import { iconButton } from './chrome-icons';
import { translateUiText } from './localization';
import type { FileCommand } from './file-commands';

export type SaveState = 'saved' | 'dirty' | 'saving' | 'saved-local';

export interface TitleBarHandlers {
	fileCommand(command: FileCommand): void;
	history(key: 'undo' | 'redo'): void;
	toggleComments(): void;
	setReadOnly(readOnly: boolean): void;
	/** The ribbon root, searched by the "Tell me" box for runnable commands. */
	ribbon(): HTMLElement | undefined;
}

export interface TitleBar {
	element: HTMLElement;
	setFileName(name: string): void;
	setSaveState(state: SaveState): void;
	setReadOnly(readOnly: boolean): void;
	setCommentsOpen(open: boolean): void;
	/** Redraws every label after the editor locale changed. */
	relocalize(): void;
}

type TitleBarElement = HTMLElement & { state: OfficeTitleBarState };

const SAVE_STATE_TEXT: Record<SaveState, string> = {
	saved: 'Saved',
	dirty: 'Unsaved changes',
	saving: 'Saving…',
	'saved-local': 'Saved to this PC',
};

/** Ribbon commands the "Tell me" box can run: every labelled, enabled button or select. */
function ribbonCommands(ribbon: HTMLElement): { label: string; control: HTMLElement }[] {
	const seen = new Set<string>();
	const commands: { label: string; control: HTMLElement }[] = [];
	for (const control of ribbon.querySelectorAll<HTMLElement>(
		'.ribbon-panel button[aria-label]:not([data-dve-hidden]), .ribbon-panel select[aria-label]:not([data-dve-hidden])',
	)) {
		const label = control.getAttribute('aria-label')!;
		if (seen.has(label)) continue;
		seen.add(label);
		commands.push({ label, control });
	}
	return commands;
}

/** Shows the ribbon tab holding `control`, then activates (button) or focuses (select) it. */
function runCommand(ribbon: HTMLElement, control: HTMLElement): void {
	const panel = control.closest<HTMLElement>('.ribbon-panel');
	const tab = panel && ribbon.querySelector<HTMLButtonElement>(`[aria-controls="${panel.id}"]`);
	tab?.click();
	if (control instanceof HTMLSelectElement) control.focus();
	else (control as HTMLButtonElement).click();
}

/** The most commands the "Tell me" list shows. */
const TELL_ME_LIMIT = 8;

/**
 * Word-style title bar: quick access toolbar, document name/state and "Tell me". The shared
 * `office-ui-title-bar` draws it from translated state; the comments toggle and view mode stay
 * docx controls placed in its `account` slot.
 */
export function createTitleBar(handlers: TitleBarHandlers): TitleBar {
	// The bar renders a search field and a switch; register them here so this works on every
	// ooxml-ui release, including ones whose defineTitleBar() does not.
	defineSearchField();
	defineSwitch();
	defineTitleBar();
	const element = document.createElement('office-ui-title-bar') as TitleBarElement;
	element.className = 'dve-titlebar';

	const actions = document.createElement('div');
	actions.className = 'dve-title-actions';
	actions.slot = 'account';
	const comments = iconButton('comment', 'Show comments');
	comments.setAttribute('aria-pressed', 'false');
	comments.addEventListener('click', () => handlers.toggleComments());
	const mode = document.createElement('select');
	mode.className = 'dve-mode-select';
	mode.setAttribute('aria-label', 'Editing mode');
	const modeOptions: [string, string][] = [
		['editing', 'Editing'],
		['viewing', 'Viewing'],
	];
	for (const [value, text] of modeOptions) {
		const option = document.createElement('option');
		option.value = value;
		option.textContent = text;
		mode.append(option);
	}
	mode.addEventListener('change', () => handlers.setReadOnly(mode.value === 'viewing'));
	actions.append(comments, mode);
	element.append(actions);

	const model = { fileName: '', saveState: 'saved' as SaveState, readOnly: false };
	/** The "Tell me" matches of the last query, by the id the bar reports back. */
	let found = new Map<string, HTMLElement>();

	const render = () => {
		const t = (text: string) => translateUiText(element, text);
		const quick = (id: string, icon: string, label: string, disabled = false) => ({
			id,
			icon,
			label: t(label),
			title: t(label),
			disabled,
		});
		const prompt = t('Tell me what you want to do');
		element.state = {
			appMark: 'W',
			fileName: model.fileName,
			status: t(SAVE_STATE_TEXT[model.saveState]),
			tone: model.saveState === 'saving' ? 'saving' : 'idle',
			quickAccess: {
				label: t('Quick access'),
				items: [
					quick('save', 'save', 'Save'),
					quick('undo', 'undo', 'Undo', model.readOnly),
					quick('redo', 'redo', 'Redo', model.readOnly),
				],
			},
			search: {
				placeholder: prompt,
				label: prompt,
				heading: prompt,
				empty: t('No matching commands'),
				match(query) {
					const ribbon = handlers.ribbon();
					const needle = query.trim().toLocaleLowerCase();
					found = new Map();
					if (!ribbon || !needle) return [];
					const matches: OfficeTitleBarCommand[] = [];
					for (const { label, control } of ribbonCommands(ribbon)) {
						if ((control as HTMLButtonElement).disabled) continue;
						if (!label.toLocaleLowerCase().includes(needle)) continue;
						found.set(label, control);
						matches.push({ id: label, label });
						if (matches.length === TELL_ME_LIMIT) break;
					}
					return matches;
				},
			},
		};
	};

	element.addEventListener('office-command', (event) => {
		const { command } = (event as CustomEvent<{ command: string }>).detail;
		if (command === 'save') handlers.fileCommand('save');
		else if (command === 'undo' || command === 'redo') handlers.history(command);
	});
	element.addEventListener('office-command-search', (event) => {
		const { command } = (event as CustomEvent<{ command?: string }>).detail;
		const control = command ? found.get(command) : undefined;
		const ribbon = handlers.ribbon();
		if (control && ribbon) runCommand(ribbon, control);
	});

	render();
	return {
		element,
		setFileName(value) {
			model.fileName = value;
			render();
		},
		setSaveState(value) {
			model.saveState = value;
			render();
		},
		setReadOnly(readOnly) {
			model.readOnly = readOnly;
			mode.value = readOnly ? 'viewing' : 'editing';
			render();
		},
		setCommentsOpen(open) {
			comments.setAttribute('aria-pressed', String(open));
		},
		relocalize: render,
	};
}
