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
import { selectedEquationPos } from './equation-tab';
import { isTabHidden, panelOf, selectTab, setTabHidden } from './ribbon-tab-api';
import type { EditorState } from 'prosemirror-state';
import { focusView } from './focus-view';
import type { RibbonAction } from './ribbon-action';
import type { EditorLocale } from './localization';

type EquationAction = Extract<RibbonAction, { type: 'equation' }>;

export interface EquationSessionHost {
	view(): EditorView | undefined;
	toolbar(): HTMLElement | undefined;
	locale(): EditorLocale;
}

/**
 * Shows the tab while an equation is selected (switching to it as the selection enters one),
 * enables its commands when the document can be edited and reflects the input format and
 * whether Normal Text is on.
 */
export function syncEquationTab(
	toolbar: HTMLElement,
	state: EditorState,
	editable: boolean,
	status: { format: 'unicode' | 'latex'; normalText: boolean },
): void {
	const panel = panelOf(toolbar, 'equation');
	if (!panel) return;
	const selected = selectedEquationPos(state) !== undefined;
	const entering = isTabHidden(panel) && selected;
	setTabHidden(toolbar, 'equation', 'contextual', !selected);
	if (entering) selectTab(toolbar, 'equation');
	for (const control of panel.querySelectorAll<HTMLButtonElement | HTMLSelectElement>(
		'button:not([data-unavailable]), select',
	))
		control.disabled = !selected || !editable;
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
				openEquationGallery(anchor, value, (chosen) => {
					if (chosen.type === 'equation') this.handle(chosen);
				});
			return;
		}
		if (!view.editable) return;
		const { box } = this;
		const attached =
			box.attachedAt !== undefined && box.attachedAt === selectedEquationPos(view.state);
		if (key === 'new') {
			if (insertNewEquation(view) !== undefined) this.sync(true);
		} else if (key === 'builtIn') {
			// An empty equation takes the built-in one; otherwise it is a new equation.
			const empty = attached && box.input.value.trim() === '' ? box.attachedAt : undefined;
			placeBuiltIn(view, value, empty);
			focusView(view);
			this.sync(true);
		} else if (key === 'convert') {
			const linear = value.endsWith('Linear');
			if (value.startsWith('all')) convertAllEquations(view, linear, this.format);
			else if (attached) box.setLinear(linear);
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
		const entering = pos !== undefined && pos !== this.lastPos;
		this.lastPos = pos;
		if (view && pos !== undefined) this.box.sync(view, pos, focus || (entering && view.hasFocus()));
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
			});
	}

	reset(): void {
		this.box.detach();
		this.lastPos = undefined;
	}
}
