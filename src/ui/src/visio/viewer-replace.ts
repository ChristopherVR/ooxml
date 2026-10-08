import {
	editErrorMessage,
	isEditCancellation,
	VISIO_TEXT_REPLACE_LIMITS,
	type VisioTextOccurrence,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import type { ViewerTextReplaceToken } from './replacement-token';
import { renderFindBar, type FindBar } from './viewer-search';

const scopes = ['selection', 'current-page', 'all-pages'] as const;
type Scope = (typeof scopes)[number];
import { sameTextOccurrence as same, renderReplaceStatus } from './viewer-replace-status';

/** Literal occurrence draft. Controller tokens retain scope through owned result navigation. */
export class ViewerReplace {
	#token: ViewerTextReplaceToken | undefined;
	#matches: readonly VisioTextOccurrence[] = [];
	#current: VisioTextOccurrence | undefined;
	#active = false;
	#navigating = false;
	#pending: { state: ViewerState; source: number; token: ViewerTextReplaceToken } | undefined;
	#request = 0;
	#summary = '';
	constructor(
		private readonly bar: FindBar,
		private readonly controller: ViewerController,
		private readonly announce: (message: string) => void,
	) {
		bar.scope = 'current-page';
		bar.scopeOptions = [
			{ value: 'selection', label: 'Selection' },
			{ value: 'current-page', label: 'Current page' },
			{ value: 'all-pages', label: 'All pages' },
		];
		bar.matchCase = true;
		bar.matchCaseDisabled = true;
		bar.replacementMaxlength = VISIO_TEXT_REPLACE_LIMITS.replacement;
	}
	wire(): () => void {
		this.#active = true;
		const Abort = this.bar.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.bar.addEventListener(
			'office-find-input',
			() => {
				if (this.bar.replaceMode) {
					this.#reset();
					this.#refresh();
				}
			},
			options,
		);
		this.bar.addEventListener(
			'office-find-options',
			(event) => {
				if (!this.bar.replaceMode) return;
				const scope = (event as CustomEvent<{ scope: string }>).detail.scope;
				if (!scopes.includes(scope as Scope)) {
					this.bar.error = 'Choose a supported search scope.';
					return;
				}
				if (this.#pending) this.#reset();
				if (this.#scope !== scope) this.#reset();
				this.#scope = scope as Scope;
				this.#summary = '';
				this.bar.error = '';
				this.#refresh();
			},
			options,
		);
		this.bar.addEventListener(
			'office-find-step',
			(event) => {
				if (this.bar.replaceMode)
					this.#step((event as CustomEvent<{ direction: 'next' | 'previous' }>).detail.direction);
			},
			options,
		);
		this.bar.addEventListener(
			'office-find-replace',
			(event) => {
				if (this.bar.replaceMode)
					void this.#replace((event as CustomEvent<{ mode: 'current' | 'all' }>).detail.mode);
			},
			options,
		);
		this.bar.addEventListener('office-find-close', () => this.#reset(), options);
		return () => {
			this.#active = false;
			this.#reset();
			events.abort();
		};
	}
	#scope: Scope = 'current-page';
	showReplace(): void {
		if (!this.#active) return;
		this.bar.replaceMode = true;
		this.bar.setAttribute('previous-label', 'Previous occurrence');
		this.bar.setAttribute('next-label', 'Next occurrence');
		this.bar.show();
		this.#refresh();
	}
	showFind(): void {
		this.#reset();
		this.bar.replaceMode = false;
		this.bar.error = '';
		this.bar.setAttribute('previous-label', 'Previous matching shape');
		this.bar.setAttribute('next-label', 'Next matching shape');
		this.controller.setSearchQuery(this.bar.value);
		renderFindBar(this.bar, this.controller.state);
		this.bar.show();
	}
	#reset(): void {
		const pending = this.#pending;
		++this.#request;
		this.#token = undefined;
		this.#current = undefined;
		this.#matches = [];
		this.#summary = '';
		this.#pending = undefined;
		this.bar.error = '';
		if (pending) this.controller.cancelTextReplace(pending.token);
	}
	#refresh(force = false): void {
		if (!this.#active || !this.bar.open || !this.bar.replaceMode || this.#pending) return;
		const state = this.controller.state;
		if (
			state.loading ||
			state.edit.busy ||
			!state.edit.sourceAvailable ||
			!state.document ||
			!this.bar.value
		) {
			this.#matches = [];
			this.#paint(state);
			return;
		}
		const request = this.#request;
		try {
			if (!force && this.#token && this.controller.isTextReplaceTokenCurrent(this.#token)) {
				this.#paint(state);
				return;
			}
			const token = this.#token ?? this.controller.captureTextReplaceToken(this.#scope);
			const matches = this.controller.getTextReplaceOccurrences(token, this.bar.value);
			if (
				!this.#active ||
				request !== this.#request ||
				!this.controller.isTextReplaceTokenCurrent(token)
			)
				return;
			this.#token = token;
			this.#matches = matches;
			if (this.#current && !matches.some((item) => same(item, this.#current!)))
				this.#current = undefined;
		} catch (error) {
			if (
				request === this.#request &&
				this.#active &&
				this.bar.replaceMode &&
				!isEditCancellation(error)
			)
				this.bar.error = editErrorMessage(error);
		}
		this.#paint(this.controller.state);
	}
	render(state: ViewerState): void {
		if (!this.bar.replaceMode) return;
		if (this.#pending) {
			const pending = this.#pending;
			if (
				state.loading ||
				!state.edit.sourceAvailable ||
				this.controller.sourceGeneration !== pending.source ||
				(state.edit.busy &&
					(state.document !== pending.state.document ||
						state.pageIndex !== pending.state.pageIndex ||
						state.selectedShapes !== pending.state.selectedShapes))
			)
				this.#reset();
		} else if ((!state.edit.sourceAvailable || state.loading || !state.document) && this.#token)
			this.#reset();
		else if (
			this.#token &&
			!this.#navigating &&
			!this.controller.isTextReplaceTokenCurrent(this.#token)
		)
			this.#reset();
		if (!this.#navigating && !this.#pending && !state.edit.busy && !this.#token) this.#refresh();
		this.#paint(state);
	}
	#paint(state: ViewerState): void {
		renderReplaceStatus(
			this.bar,
			state,
			this.#matches,
			this.#current,
			!!this.#pending,
			this.#summary,
		);
	}
	#step(direction: 'next' | 'previous'): void {
		if (this.#pending || !this.#token || !this.#matches.length) return;
		const token = this.#token,
			request = this.#request;
		const index = this.#current
			? this.#matches.findIndex((item) => same(item, this.#current!))
			: -1;
		const next =
			direction === 'next'
				? (index + 1) % this.#matches.length
				: index <= 0
					? this.#matches.length - 1
					: index - 1;
		this.#navigating = true;
		try {
			const current = this.#matches[next]!;
			const fresh = this.controller.selectTextReplaceOccurrence(token, current, this.bar.value);
			if (
				!this.#active ||
				request !== this.#request ||
				!this.controller.isTextReplaceTokenCurrent(fresh)
			)
				return;
			this.#token = fresh;
			this.#current = current;
			this.#summary = '';
			this.bar.error = '';
		} catch (error) {
			this.#reset();
			if (!isEditCancellation(error)) this.bar.error = editErrorMessage(error);
		} finally {
			this.#navigating = false;
		}
		if (!this.#token) this.#refresh();
		this.#paint(this.controller.state);
	}
	async #replace(mode: 'current' | 'all'): Promise<void> {
		if (this.#pending || !this.#token || !this.#matches.length || !this.#active) return;
		const token = this.#token,
			request = ++this.#request;
		this.bar.error = '';
		try {
			const plan = this.controller.planTextReplacement(token, {
				query: this.bar.value,
				replacement: this.bar.replacement,
				mode,
				...(this.#current ? { current: this.#current } : {}),
			});
			if (request !== this.#request || !this.controller.isTextReplaceTokenCurrent(token)) return;
			if (!plan.replacementCount) return;
			this.#pending = {
				state: this.controller.state,
				source: this.controller.sourceGeneration,
				token,
			};
			const fresh = await this.controller.applyTextReplacePlan(plan, token);
			if (
				!this.#active ||
				request !== this.#request ||
				!this.bar.replaceMode ||
				!this.controller.isTextReplaceTokenCurrent(fresh)
			)
				return;
			this.#pending = undefined;
			this.#token = fresh;
			this.#current = plan.nextOccurrence;
			this.#refresh(true);
			if (!this.#owns(request, fresh)) return;
			this.#summary = `Replaced ${plan.replacementCount} occurrence${plan.replacementCount === 1 ? '' : 's'}.`;
			this.#paint(this.controller.state);
			if (this.#owns(request, fresh)) this.announce(this.#summary);
		} catch (error) {
			if (request !== this.#request || !this.#active) return;
			this.#pending = undefined;
			this.#token = undefined;
			this.#refresh();
			if (!isEditCancellation(error)) this.bar.error = editErrorMessage(error);
		}
	}
	#owns(request: number, token: ViewerTextReplaceToken): boolean {
		return (
			this.#active &&
			request === this.#request &&
			this.bar.open &&
			this.bar.replaceMode &&
			this.#token === token &&
			this.controller.isTextReplaceTokenCurrent(token)
		);
	}
}
