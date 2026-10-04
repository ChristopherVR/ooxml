/**
 * Context menu for the document surface. Opens on right click, Shift+F10 or the ContextMenu key.
 * What it offers comes from context-menu-items.ts; this module owns the DOM, keyboard navigation,
 * dismissal and viewport clamping in the shared `office-ui-context-menu` element; this module
 * wires it to the editor.
 */
import { defineContextMenu } from 'ooxml-ui/controls';
import type { OfficeMenuItem, OfficeMenuState } from 'ooxml-ui/controls';
import type { EditorView } from 'prosemirror-view';
import { clipboardReadAllowed, copyOrCut, pasteClipboard } from './context-menu-actions';
import { computeMenuItems, menuStateFromView, type MenuItem } from './context-menu-items';
import { removeLink } from './link-commands';
import { translate, type EditorLocale } from './localization';
import type { RibbonAction } from './ribbon-action';

export interface ContextMenuHost {
	/** Element the menu is appended to (inside the shadow root, so theme tokens apply). */
	container: HTMLElement;
	/** Element whose `contextmenu` events open the menu. */
	target: HTMLElement;
	view(): EditorView | undefined;
	/** Returns focus to the document surface (which may be non-editable). */
	focusDocument(): void;
	readOnly(): boolean;
	locale(): EditorLocale;
	run(action: RibbonAction): void;
	warn(message: string): void;
}

export interface ContextMenu {
	readonly isOpen: boolean;
	readonly element: HTMLElement | undefined;
	open(x: number, y: number): void;
	/** Opens at the caret, for Shift+F10 and the ContextMenu key. */
	openAtCaret(): void;
	close(restoreFocus?: boolean): void;
}

/** Browsers may also fire `contextmenu` for the key that we already handled; ignore that echo. */
const KEY_ECHO_MS = 500;

export function createContextMenu(host: ContextMenuHost): ContextMenu {
	let menu: (HTMLElement & { state: OfficeMenuState }) | undefined;
	let disposeDismiss: (() => void) | undefined;
	let keyOpenedAt = -Infinity;

	const close = (restoreFocus = true) => {
		if (!menu) return;
		menu.remove();
		menu = undefined;
		disposeDismiss?.();
		disposeDismiss = undefined;
		if (restoreFocus) host.focusDocument();
	};

	const execute = async (item: MenuItem, view: EditorView) => {
		const { command } = item;
		if (command.kind === 'ribbon') host.run(command.action);
		else if (command.kind === 'removeLink') {
			removeLink(view);
			view.focus();
		} else if (command.op === 'paste') {
			if (!(await pasteClipboard(view)))
				host.warn(translate(host.locale(), 'menu.clipboardDenied'));
		} else if (!(await copyOrCut(view, command.op)))
			host.warn(translate(host.locale(), 'menu.clipboardDenied'));
	};

	const rows = (items: MenuItem[], locale: EditorLocale): OfficeMenuItem[] =>
		items.map((item) => ({
			id: item.id,
			label: translate(locale, item.label),
			disabled: item.disabled,
			separatorBefore: item.separatorBefore,
			...(item.disabled && item.reason ? { title: translate(locale, item.reason) } : {}),
		}));

	const watchDismissal = () => {
		const dismiss = () => close(false);
		window.addEventListener('scroll', dismiss, true);
		// `scroll` is not composed, so the shadow-tree canvas needs its own listener.
		host.target.addEventListener('scroll', dismiss);
		window.addEventListener('resize', dismiss);
		window.addEventListener('blur', dismiss);
		return () => {
			window.removeEventListener('scroll', dismiss, true);
			host.target.removeEventListener('scroll', dismiss);
			window.removeEventListener('resize', dismiss);
			window.removeEventListener('blur', dismiss);
		};
	};

	const open = (x: number, y: number) => {
		const view = host.view();
		if (!view) return;
		close(false);
		defineContextMenu();
		const locale = host.locale();
		const items = computeMenuItems(menuStateFromView(view, host.readOnly()));
		const element = document.createElement('office-ui-context-menu') as HTMLElement & {
			state: OfficeMenuState;
		};
		const state: OfficeMenuState = {
			x,
			y,
			label: translate(locale, 'menu.context'),
			items: rows(items, locale),
		};
		element.state = state;
		element.addEventListener('office-menu-request', (event) => {
			const id = (event as CustomEvent<{ id: string }>).detail.id;
			const item = items.find((candidate) => candidate.id === id);
			const current = host.view();
			if (!item || !current) return;
			close();
			void execute(item, current);
		});
		// Escape and Tab restore focus to the document; an outside click leaves it where it landed.
		element.addEventListener('office-menu-close', (event) =>
			close((event as CustomEvent<{ reason: string }>).detail.reason !== 'outside'),
		);
		menu = element;
		host.container.append(element);
		disposeDismiss = watchDismissal();
		if (items.some((item) => item.id === 'paste' && !item.disabled))
			void clipboardReadAllowed().then((allowed) => {
				if (allowed || menu !== element) return;
				const denied = translate(locale, 'menu.clipboardDenied');
				element.state = {
					...state,
					items: state.items.map((row) =>
						row.id === 'paste' ? { ...row, disabled: true, title: denied } : row,
					),
				};
			});
	};

	host.target.addEventListener('contextmenu', (event) => {
		if (!(event.target as Element).closest?.('.ProseMirror')) return;
		event.preventDefault();
		if (performance.now() - keyOpenedAt < KEY_ECHO_MS) return;
		const { clientX, clientY } = event;
		// Let the browser and ProseMirror move the caret to the click before reading the selection.
		setTimeout(() => open(clientX, clientY), 0);
	});

	return {
		get isOpen() {
			return Boolean(menu);
		},
		get element() {
			return menu;
		},
		open,
		openAtCaret() {
			const view = host.view();
			if (!view) return;
			keyOpenedAt = performance.now();
			const caret = view.coordsAtPos(view.state.selection.head);
			open(caret.left, caret.bottom);
		},
		close,
	};
}
