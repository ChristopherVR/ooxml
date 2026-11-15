/**
 * The Paste Options button Word shows at the end of a pasted range, and its menu (the shared
 * `office-ui-context-menu`): Keep Source Formatting, Merge Formatting, Keep Text Only. Ctrl on
 * its own opens the menu while the button shows; Escape withdraws it.
 */
import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { defineContextMenu, type OfficeMenuItem, type OfficeMenuState } from '../controls';
import { translate, type EditorLocale, type LocalizationKey } from './localization';
import {
	applyPasteOption,
	dismissPasteOptions,
	pastedRange,
	type PasteOption,
	type PastedRange,
} from './paste-options';
import { ribbonIcon } from './ribbon-icons';

export interface PasteOptionsButtonHost {
	/** Where the button and its menu are appended (inside the shadow root, for the theme). */
	container(): HTMLElement | undefined;
	locale(): EditorLocale;
}

const OPTIONS: Array<[PasteOption, LocalizationKey]> = [
	['source', 'Keep Source Formatting'],
	['merge', 'Merge Formatting'],
	['text', 'Keep Text Only'],
];

type MenuElement = HTMLElement & { state: OfficeMenuState };

/** The menu rows for `range`: the current choice is checked, HTML choices need HTML. */
export function pasteOptionItems(range: PastedRange, locale: EditorLocale): OfficeMenuItem[] {
	return OPTIONS.map(([id, label]) => ({
		id,
		label: translate(locale, label),
		checked: range.option === id,
		disabled: id !== 'text' && !range.payload.html,
	}));
}

class PasteOptionsView {
	private button: HTMLButtonElement | undefined;
	private menu: MenuElement | undefined;
	private shown: PastedRange | null = null;
	private frame = 0;
	private ctrlAlone = false;
	private readonly reposition = () => this.place();

	constructor(
		private readonly view: EditorView,
		private readonly host: PasteOptionsButtonHost,
	) {
		view.dom.addEventListener('keydown', this.keydown);
		view.dom.addEventListener('keyup', this.keyup);
		window.addEventListener('scroll', this.reposition, true);
		window.addEventListener('resize', this.reposition);
	}

	private readonly keydown = (event: KeyboardEvent) => {
		this.ctrlAlone = event.key === 'Control' && !event.altKey && !event.shiftKey;
		if (event.key === 'Escape' && this.shown && !this.menu && dismissPasteOptions(this.view))
			event.preventDefault();
	};

	private readonly keyup = (event: KeyboardEvent) => {
		if (event.key === 'Control' && this.ctrlAlone && this.shown) this.openMenu();
		this.ctrlAlone = false;
	};

	update(): void {
		const range = pastedRange(this.view);
		if (range === this.shown) return;
		this.shown = range;
		this.closeMenu();
		if (!range) {
			this.button?.remove();
			this.button = undefined;
			return;
		}
		const container = this.host.container();
		if (!container) return;
		if (!this.button) {
			const button = document.createElement('button');
			button.type = 'button';
			button.className = 'dve-paste-options';
			button.setAttribute('aria-haspopup', 'menu');
			button.append(ribbonIcon('paste', 16));
			// Clicking must not move the caret or blur the document before the menu opens.
			button.addEventListener('mousedown', (event) => event.preventDefault());
			button.addEventListener('click', () => this.openMenu());
			this.button = button;
		}
		const label = translate(this.host.locale(), 'Paste Options (Ctrl)');
		this.button.title = label;
		this.button.setAttribute('aria-label', label);
		container.append(this.button);
		this.place();
	}

	/** Puts the button just below and after the end of the pasted range, once layout settles
	 * (the print layout redraws the visible caret, which the editor's own text follows, after a
	 * transaction). */
	private place(): void {
		cancelAnimationFrame(this.frame);
		this.frame = requestAnimationFrame(() => this.placeNow());
	}

	private placeNow(): void {
		const range = this.shown;
		const button = this.button;
		if (!range || !button || this.view.isDestroyed) return;
		try {
			const end = this.view.coordsAtPos(Math.min(range.to, this.view.state.doc.content.size));
			button.style.left = `${Math.round(end.right + 2)}px`;
			button.style.top = `${Math.round(end.bottom + 2)}px`;
		} catch {
			// No geometry (hidden view): leave the button where it was.
		}
	}

	private openMenu(): void {
		const range = this.shown;
		const container = this.host.container();
		if (!range || !container || !this.button) return;
		this.closeMenu();
		defineContextMenu();
		const rect = this.button.getBoundingClientRect();
		const locale = this.host.locale();
		const menu = document.createElement('office-ui-context-menu') as MenuElement;
		menu.state = {
			x: rect.left,
			y: rect.bottom,
			label: translate(locale, 'Paste Options (Ctrl)'),
			items: pasteOptionItems(range, locale),
		};
		menu.addEventListener('office-menu-request', (event) => {
			const id = (event as CustomEvent<{ id: string }>).detail.id as PasteOption;
			this.closeMenu();
			applyPasteOption(this.view, id);
			this.view.focus();
		});
		menu.addEventListener('office-menu-close', () => {
			this.closeMenu();
			this.view.focus();
		});
		this.menu = menu;
		container.append(menu);
	}

	private closeMenu(): void {
		this.menu?.remove();
		this.menu = undefined;
	}

	destroy(): void {
		cancelAnimationFrame(this.frame);
		this.closeMenu();
		this.button?.remove();
		this.view.dom.removeEventListener('keydown', this.keydown);
		this.view.dom.removeEventListener('keyup', this.keyup);
		window.removeEventListener('scroll', this.reposition, true);
		window.removeEventListener('resize', this.reposition);
	}
}

/** Draws the Paste Options button for the body editor. */
export function pasteOptionsButtonPlugin(host: PasteOptionsButtonHost): Plugin {
	return new Plugin({ view: (view) => new PasteOptionsView(view, host) });
}
