// Equations in the Word editor: Insert New Equation (Alt+=), the contextual Equation tab and the
// in-place equation editor. The editor, the tab's commands and the galleries load on first use
// (`equation-session.ts`); until then this only keeps the tab hidden.
import { keymap } from 'prosemirror-keymap';
import type { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { EquationSession } from './equation-session';
import { selectedEquationPos } from './equation-tab';
import type { RibbonAction } from './ribbon-action';
import { panelOf, setTabHidden } from './ribbon-tab-api';
import type { EditorLocale } from './localization';

export interface EquationControllerHost {
	view(): EditorView | undefined;
	toolbar(): HTMLElement | undefined;
	reportError(error: Error): void;
}

export class EquationController {
	private locale: EditorLocale = 'en';
	private session: EquationSession | undefined;
	private loading: Promise<EquationSession> | undefined;

	constructor(private readonly host: EquationControllerHost) {}

	setLocale(locale: EditorLocale): void {
		this.locale = locale;
	}

	private load(): Promise<EquationSession> {
		this.loading ??= Promise.all([
			import('./equation-session'),
			import('./equation-kit').then((kit) => kit.loadEquationKit()),
		]).then(
			([{ EquationSession }]) =>
				(this.session = new EquationSession({
					view: () => this.host.view(),
					toolbar: () => this.host.toolbar(),
					locale: () => this.locale,
				})),
		);
		return this.loading;
	}

	private fail = (cause: unknown) =>
		this.host.reportError(cause instanceof Error ? cause : new Error(String(cause)));

	/** Alt+= inserts a new equation, as in Word. */
	plugins(): Plugin[] {
		return [
			keymap({
				'Alt-=': (_state, _dispatch, view) =>
					Boolean(view?.editable) && this.handle({ type: 'equation', key: 'new' }),
			}),
		];
	}

	handle(action: RibbonAction): boolean {
		if (action.type !== 'equation') return false;
		void this.load().then((session) => session.handle(action), this.fail);
		return true;
	}

	/** Shows the Equation tab and the equation editor while an equation is selected. */
	sync(): void {
		if (this.session) return this.session.sync();
		const view = this.host.view();
		if (view && selectedEquationPos(view.state) !== undefined)
			void this.load().then((session) => session.sync(), this.fail);
		else {
			const toolbar = this.host.toolbar();
			if (toolbar && panelOf(toolbar, 'equation'))
				setTabHidden(toolbar, 'equation', 'contextual', true);
		}
	}

	reset(): void {
		this.session?.reset();
	}
}
