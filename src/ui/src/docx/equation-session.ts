// The Equation tab and the in-place equation editor once equations are in use (loaded by
// `EquationController` on first use): it opens the editor on the selected equation, keeps the
// tab's state and runs the tab's commands.
import type { EditorView } from 'prosemirror-view';
import type { EquationInputFormat } from 'ooxml-core/math';
import {
	convertAllEquations,
	EquationEditingBox,
	insertNewEquation,
	placeBuiltIn,
	structureById,
} from './equation-editing';
import { openEquationGallery } from './equation-galleries';
import { sameEquation, selectedEquationPos } from './equation-tab';
import {
	isTabHidden,
	panelOf,
	selectTab,
	setTabHidden,
	type RibbonElement,
} from './ribbon-tab-api';
import type { EditorState } from 'prosemirror-state';
import { emit } from './events';
import { focusView } from './focus-view';
import type { RibbonAction } from './ribbon-action';
import { translate, type EditorLocale, type LocalizationKey } from './localization';
import { inkMathRecognizer } from '../ink/math-recognizer';

type EquationAction = Extract<RibbonAction, { type: 'equation' }>;

export interface EquationSessionHost {
	view(): EditorView | undefined;
	toolbar(): HTMLElement | undefined;
	locale(): EditorLocale;
}

/**
 * Shows the tab while an equation is selected (switching to it as the selection enters one),
 * enables its commands when the document and the equation can be edited and reflects the input
 * format and whether Normal Text is on.
 */
export function syncEquationTab(
	toolbar: HTMLElement,
	state: EditorState,
	editable: boolean,
	status: {
		format: 'unicode' | 'latex';
		normalText: boolean;
		canEdit: boolean;
		/** Why Ink Equation cannot run (no recognizer), translated; undefined when it can. */
		inkUnavailable?: string | undefined;
		/** Ink Equation's tooltip while it can run, translated. */
		inkTitle?: string;
	},
): void {
	const panel = panelOf(toolbar, 'equation');
	if (!panel) return;
	const selected = selectedEquationPos(state) !== undefined;
	const entering = isTabHidden(panel) && selected;
	setTabHidden(toolbar, 'equation', 'contextual', !selected);
	if (entering) {
		selectTab(toolbar, 'equation');
		// A narrow ribbon scrolls its tab strip so the new tab is in view.
		(toolbar as RibbonElement)
			.tabButton('equation')
			?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
	}
	for (const control of panel.querySelectorAll<HTMLButtonElement | HTMLSelectElement>(
		'button:not([data-unavailable]), select',
	))
		// A read-only equation still takes the gallery and Convert > All, which leave it alone.
		control.disabled =
			!selected || !editable || (!status.canEdit && !control.closest('[data-equation-any]'));
	const ink = panel.querySelector<HTMLButtonElement>('[data-equation-ink]');
	if (ink) {
		if (status.inkUnavailable) ink.disabled = true;
		ink.title = status.inkUnavailable ?? status.inkTitle ?? ink.title;
	}
	for (const button of panel.querySelectorAll<HTMLElement>('[data-equation-format]'))
		button.setAttribute('aria-pressed', String(button.dataset.equationFormat === status.format));
	panel
		.querySelector('.dve-equation-normal-text')
		?.setAttribute('aria-pressed', String(status.normalText));
}

export class EquationSession {
	/** The input format linear text is typed in; Word starts in Unicode. */
	format: EquationInputFormat = 'unicode';
	readonly box: EquationEditingBox;
	/** The equation the selection was on at the last sync, to focus the editor on entering one. */
	private lastPos: number | undefined;

	constructor(private readonly host: EquationSessionHost) {
		this.box = new EquationEditingBox({
			view: () => host.view(),
			toolbar: () => host.toolbar(),
			locale: () => host.locale(),
			format: () => this.format,
			setFormat: (format) => (this.format = format),
			refresh: () => this.syncTab(),
		});
	}

	handle(action: EquationAction): void {
		const view = this.host.view();
		if (!view) return;
		const { key, value = '' } = action;
		if (key === 'gallery') {
			const anchor = this.host
				.toolbar()
				?.querySelector<HTMLElement>(`[data-equation-gallery="${value}"]`);
			if (anchor)
				openEquationGallery(anchor, value.replace(/^insert-/, ''), {
					choose: (chosen) => {
						const toolbar = this.host.toolbar();
						if (chosen.type === 'equation') this.handle(chosen);
						else if (toolbar) emit(toolbar, 'ribbon-action', chosen);
					},
					// Escape gives the caret back to the equation being edited, as Word does.
					dismissed: () => {
						if (this.box.attachedAt !== undefined) this.box.input.focus({ preventScroll: true });
					},
				});
			return;
		}
		if (!view.editable) return;
		const { box } = this;
		const attached =
			box.attachedAt !== undefined && box.attachedAt === selectedEquationPos(view.state);
		if (key === 'ink') {
			const toolbar = this.host.toolbar();
			const root = toolbar?.getRootNode();
			void import('./ink-equation').then(({ openWordInkEquation }) =>
				openWordInkEquation({
					host: root instanceof ShadowRoot ? root : (toolbar ?? document.body),
					view: () => this.host.view(),
					locale: () => this.host.locale(),
					heading: 'Ink Equation',
				}),
			);
			return;
		}
		if (key === 'new') {
			if (insertNewEquation(view) !== undefined) this.sync(true);
		} else if (key === 'builtIn') {
			// An empty equation takes the built-in one; otherwise it is a new equation.
			// Only an editable equation that holds nothing is filled; a read-only one is never replaced,
			// nor a paragraph of several equations for one empty line of it.
			const empty =
				attached && box.editable && !box.inPart && box.input.value.trim() === ''
					? box.attachedAt
					: undefined;
			placeBuiltIn(view, value, empty);
			focusView(view);
			this.sync(true);
		} else if (key === 'convert') {
			const linear = value.endsWith('Linear');
			if (value.startsWith('all')) convertAllEquations(view, linear);
			else if (attached) box.setLinear(linear);
			// The caret stays in the equation, so undo and typing go on there.
			if (attached) this.sync(true);
		} else if (!attached) {
			return;
		} else if (key === 'structure') {
			const text = structureById(value, this.format);
			if (text) box.insert(text);
		} else if (key === 'symbol') box.insert(value);
		else if (key === 'normalText') box.toggleNormalText();
		else if (key === 'format') box.setFormat(value === 'latex' ? 'latex' : 'unicode');
		this.syncTab();
	}

	/** Shows the Equation tab and the equation editor while an equation is selected. */
	sync(focus = false): void {
		const view = this.host.view();
		const pos = view ? selectedEquationPos(view.state) : undefined;
		const same =
			view !== undefined && pos !== undefined && (sameEquation(view.state) || pos === this.lastPos);
		this.lastPos = pos;
		if (view && pos !== undefined)
			this.box.sync(view, pos, focus || (!same && view.hasFocus()), same);
		else this.box.detach();
		this.syncTab();
	}

	private syncTab(): void {
		const view = this.host.view();
		const toolbar = this.host.toolbar();
		if (view && toolbar)
			syncEquationTab(toolbar, view.state, view.editable, {
				format: this.format,
				normalText: this.box.normalText,
				canEdit: this.box.editable,
				inkUnavailable: inkMathRecognizer()
					? undefined
					: translate(
							this.host.locale(),
							'Ink Equation needs handwriting recognition, which this editor does not have.' as LocalizationKey,
						),
				inkTitle: translate(this.host.locale(), 'Ink Equation' as LocalizationKey),
			});
	}

	reset(): void {
		this.box.detach();
		this.lastPos = undefined;
	}
}
