import { canEditDiagram } from './controller';
import { checkVisioAlternativeText, type VisioShape } from 'ooxml-core/visio';
import { editErrorMessage, isEditCancellation, visioSelectionIsOnPage } from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { selectedShape } from './shape-inspector';
import { InsertDialog } from './viewer-insert-dialog';

/** Review accessibility tools, using native ShapeSheet metadata and the normal edit history. */
export class ViewerAccessibility {
	readonly alt: InsertDialog;
	readonly report: InsertDialog;
	readonly results: HTMLElement;
	#target: { pageId: string; shapeId?: string; generation: number } | undefined;
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly announce: (text: string) => void,
	) {
		this.alt = new InsertDialog(
			root,
			'alt-text-dialog',
			'Alt Text',
			[
				{ name: 'alt-title', label: 'Title' },
				{ name: 'alt-description', label: 'Description', multiline: true },
			],
			['OK', 'Cancel'],
			(button) => void this.#apply(button),
		);
		this.report = new InsertDialog(
			root,
			'accessibility-dialog',
			'Accessibility Checker',
			[],
			['Close'],
			() => this.report.dialog.close(),
		);
		const scope = root.ownerDocument.createElement('p');
		scope.textContent =
			'Checks visible graphics without text or an alternative. Review reading order and color contrast separately.';
		this.results = root.ownerDocument.createElement('div');
		this.results.setAttribute('aria-live', 'polite');
		this.report.dialog.prepend(scope, this.results);
	}
	#shape(state: ViewerState): VisioShape | undefined {
		const page = state.document?.pages[state.pageIndex];
		if (
			!page ||
			state.selectedShapes.length !== 1 ||
			!state.selectedShape ||
			!visioSelectionIsOnPage(state.selectedShape, page.id)
		)
			return;
		return selectedShape(state.document, state.selectedShape, state.pageIndex);
	}
	edit(pageOnly = false): void {
		const state = this.controller.state;
		const page = state.document?.pages[state.pageIndex];
		const shape = pageOnly ? page : this.#shape(state);
		if (!canEditDiagram(state) || state.edit.busy || state.loading || !page || !shape) return;
		this.#target = {
			pageId: page.id,
			...(pageOnly ? {} : { shapeId: shape.id }),
			generation: this.controller.documentGeneration,
		};
		this.alt.fields.get('alt-title')!.value = shape.accessibilityTitle ?? '';
		this.alt.fields.get('alt-description')!.value = shape.alternativeText ?? '';
		this.alt.error.textContent = '';
		this.alt.busy(false);
		this.alt.dialog.setAttribute('heading', pageOnly ? 'Page Alt Text' : 'Alt Text');
		this.alt.dialog.show();
		this.alt.fields.get('alt-description')!.focus();
	}
	check(): void {
		const model = this.controller.state.document;
		if (!model) return;
		const issues = checkVisioAlternativeText(model);
		this.results.replaceChildren();
		const summary = this.root.ownerDocument.createElement('p');
		summary.textContent = issues.length
			? issues.length === 1
				? '1 graphic needs alternative text.'
				: `${issues.length} graphics need alternative text.`
			: 'No missing text alternatives found.';
		this.results.append(summary);
		for (const issue of issues) {
			const button = this.root.ownerDocument.createElement('button');
			button.type = 'button';
			button.textContent = `${model.pages.find((page) => page.id === issue.pageId)?.name}: ${issue.shapeName || `Shape ${issue.shapeId}`}`;
			button.addEventListener('click', () => {
				if (this.controller.state.document !== model) return;
				const index = model.pages.findIndex((page) => page.id === issue.pageId);
				this.controller.setPage(index);
				this.controller.selectShape({
					id: issue.shapeId,
					name: issue.shapeName,
					pageId: issue.pageId,
				});
				this.report.dialog.close();
				this.edit();
			});
			this.results.append(button);
		}
		this.report.dialog.show();
	}
	async #apply(button: string): Promise<void> {
		if (button === 'Cancel') return this.alt.dialog.close();
		const target = this.#target;
		if (
			!target ||
			!this.alt.dialog.open ||
			target.generation !== this.controller.documentGeneration
		)
			return;
		try {
			const values = {
				pageId: target.pageId,
				title: this.alt.value('alt-title'),
				description: this.alt.value('alt-description'),
			};
			await this.controller.applyEdits([
				target.shapeId === undefined
					? { type: 'set-page-alt-text', ...values }
					: { type: 'set-shape-alt-text', shapeId: target.shapeId, ...values },
			]);
			if (this.#target === target) this.alt.dialog.close();
			this.announce('Saved alternative text.');
		} catch (error) {
			if (this.alt.dialog.open && this.#target === target && !isEditCancellation(error))
				this.alt.error.textContent = editErrorMessage(error);
		}
	}
	close(): void {
		this.alt.dialog.close();
		this.report.dialog.close();
	}
	render(state: ViewerState): void {
		const editable = canEditDiagram(state) && !state.loading && !state.edit.busy;
		const alt = this.root.querySelector<HTMLElement & { disabled: boolean }>(
			'[command="alt-text"]',
		);
		if (alt) alt.disabled = !editable || !this.#shape(state);
		const pageAlt = this.root.querySelector<HTMLElement & { disabled: boolean }>(
			'[command="page-alt-text"]',
		);
		if (pageAlt) pageAlt.disabled = !editable || !state.document?.pages[state.pageIndex];
		const check = this.root.querySelector<HTMLElement & { disabled: boolean }>(
			'[command="check-accessibility"]',
		);
		if (check) check.disabled = !state.document || state.loading;
		if (this.alt.dialog.open) {
			if (
				!canEditDiagram(state) ||
				state.loading ||
				this.#target?.generation !== this.controller.documentGeneration
			)
				this.alt.dialog.close();
			else this.alt.busy(state.edit.busy);
		}
		if (state.loading || !state.document) this.report.dialog.close();
	}
}
