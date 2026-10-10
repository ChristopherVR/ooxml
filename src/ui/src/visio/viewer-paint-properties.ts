import type { VisioShapeFormatEdit } from 'ooxml-core/visio';
import {
	editErrorMessage,
	isEditCancellation,
	visioSelectionIsOnPage,
	visioShapeFormattingState,
	visioFormatTargetShape,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { createPaintDialog, paintFields, type PaintField } from './viewer-paint-dialog';

type Patch = Omit<VisioShapeFormatEdit, 'type' | 'pageId' | 'shapeId'>;
type Context = { state: ViewerState; source: number; document: number };

/** Stable selection draft; one changed-fields-only worker transaction commits the dialog. */
export class ViewerPaintProperties {
	readonly #view;
	#context: Context | undefined;
	#initial = new Map<PaintField, string>();
	#pending = false;
	#request = 0;
	#active = false;
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly announce: (message: string) => void,
	) {
		this.#view = createPaintDialog(root.ownerDocument);
		root.append(this.#view.dialog);
	}
	wire(): () => void {
		this.#active = true;
		const Abort = this.#view.dialog.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.#view.apply.addEventListener('office-command', () => void this.#apply(), options);
		this.#view.cancel.addEventListener('office-command', () => this.close(), options);
		this.#view.dialog.addEventListener('office-dialog-close', () => this.#clear(), options);
		// The shared dialog consumes Escape/Tab. Other editing keys stay in its fields.
		this.#view.dialog.addEventListener(
			'keydown',
			(event) => {
				if (event.key !== 'Escape' && event.key !== 'Tab') event.stopPropagation();
			},
			options,
		);
		return () => {
			this.#active = false;
			this.close();
			events.abort();
			this.#view.dialog.remove();
		};
	}
	#selection(state: ViewerState) {
		const page = state.document?.pages[state.pageIndex];
		if (
			!page ||
			state.loading ||
			!state.edit.sourceAvailable ||
			!state.selectedShapes.length ||
			!state.selectedShapes.every((item) => visioSelectionIsOnPage(item, page.id))
		)
			return;
		const shapes = state.selectedShapes.map((item) => visioFormatTargetShape(page, item.id));
		if (shapes.some((shape) => !shape)) return;
		return { page, shapes: shapes.filter((shape) => !!shape) };
	}
	#current(context: Context): boolean {
		const state = this.controller.state;
		return (
			this.#active &&
			state.document === context.state.document &&
			state.pageIndex === context.state.pageIndex &&
			state.selectedShapes === context.state.selectedShapes &&
			!state.loading &&
			state.edit.sourceAvailable &&
			this.controller.sourceGeneration === context.source &&
			this.controller.documentGeneration === context.document
		);
	}
	show(): void {
		const state = this.controller.state;
		if (
			!this.#active ||
			state.edit.busy ||
			!state.document ||
			state.loading ||
			!state.edit.sourceAvailable
		)
			return;
		const context = {
			state,
			source: this.controller.sourceGeneration,
			document: this.controller.documentGeneration,
		};
		const selection = this.#selection(state);
		if (!selection) return;
		const common = visioShapeFormattingState(selection.shapes);
		if (!this.#current(context)) return;
		this.close();
		if (!this.#current(context)) return;
		this.#context = context;
		this.#view.error.textContent = '';
		const values = {
			fillPattern: common.fillPatternIndex,
			linePattern: common.linePattern,
			fillBackgroundColor: common.fillBackgroundColor,
			fillTransparency: common.fillTransparency,
			lineTransparency: common.lineTransparency,
		};
		for (const field of paintFields) {
			const input = this.#view.fields.get(field)!;
			if ('options' in input) {
				const length = field === 'fillPattern' ? 25 : 24;
				const current = values[field];
				input.options = [
					{ value: '', label: 'Mixed or unavailable', disabled: true },
					...(typeof current === 'number' && (current < 0 || current >= length)
						? [{ value: String(current), label: `Current pattern ${current}`, disabled: true }]
						: []),
					...Array.from({ length }, (_, index) => ({
						value: String(index),
						label:
							index === 0
								? field === 'fillPattern'
									? 'No Fill'
									: 'No Line'
								: index === 1
									? 'Solid'
									: `Pattern ${index}`,
					})),
				];
			}
			input.value = values[field] === undefined ? '' : String(values[field]);
			if (input instanceof HTMLInputElement && !input.value)
				input.placeholder = 'Mixed or unavailable';
			input.disabled = false;
			this.#initial.set(field, input.value);
		}
		this.#view.apply.disabled = false;
		if (!this.#current(context)) {
			this.close();
			return;
		}
		this.#view.dialog.show();
	}
	close(): void {
		const context = this.#context;
		this.#view.dialog.close();
		if (this.#context === context) this.#clear();
	}
	#clear(): void {
		const context = this.#context;
		const pending = this.#pending;
		++this.#request;
		this.#pending = false;
		this.#context = undefined;
		this.#initial.clear();
		if (
			pending &&
			context &&
			this.controller.state.document === context.state.document &&
			this.controller.sourceGeneration === context.source &&
			this.controller.state.edit.busy
		)
			this.controller.cancelEdit();
	}
	render(state: ViewerState): void {
		const context = this.#context;
		if (context && (!this.#pending || state.edit.busy) && !this.#current(context)) this.close();
		for (const id of ['fill-options', 'line-options', 'ctx-format']) {
			const button = this.root.querySelector<HTMLElement & { disabled: boolean }>(
				`[command="${id}"]`,
			);
			if (button) button.disabled = state.edit.busy || !this.#selection(state);
		}
	}
	async #apply(): Promise<void> {
		const context = this.#context;
		if (!context || this.#pending || this.controller.state.edit.busy || !this.#current(context))
			return;
		const request = ++this.#request;
		try {
			const patch: Patch = {};
			for (const field of paintFields) {
				const value = this.#view.fields.get(field)!.value;
				if (value === this.#initial.get(field)) continue;
				if (field === 'fillBackgroundColor') {
					if (!/^#[0-9a-f]{6}$/i.test(value))
						throw new Error('Enter a background color as #RRGGBB.');
					patch.fillBackgroundColor = value;
				} else {
					const number = Number(value);
					const maximum = field === 'fillPattern' ? 24 : field === 'linePattern' ? 23 : 100;
					if (
						!value ||
						!Number.isFinite(number) ||
						number < 0 ||
						number > maximum ||
						(field.endsWith('Pattern') && !Number.isInteger(number))
					)
						throw new Error(
							`Enter a valid ${field.endsWith('Pattern') ? 'pattern' : 'transparency from 0 to 100%'}.`,
						);
					patch[field] = number;
				}
			}
			if (!Object.keys(patch).length) {
				this.close();
				return;
			}
			const selection = this.#selection(context.state);
			if (!selection || !this.#current(context)) {
				this.close();
				return;
			}
			const commands = selection.shapes.map((shape) => ({
				type: 'format-shape' as const,
				pageId: selection.page.id,
				shapeId: shape.id,
				...patch,
			}));
			if (!this.#current(context)) {
				this.close();
				return;
			}
			this.#pending = true;
			this.#view.apply.disabled = true;
			for (const input of this.#view.fields.values()) input.disabled = true;
			await this.controller.applySelectionEdits(commands);
			if (request !== this.#request) return;
			this.#pending = false;
			const accepted = this.controller.state;
			const source = this.controller.sourceGeneration;
			const document = this.controller.documentGeneration;
			this.close();
			const closedRequest = this.#request;
			await this.#view.dialog.updateComplete;
			if (!this.#active || this.#request !== closedRequest) return;
			const current = this.controller.state;
			if (
				!this.#view.dialog.open &&
				!this.#context &&
				current.document === accepted.document &&
				current.pageIndex === accepted.pageIndex &&
				current.selectedShapes === accepted.selectedShapes &&
				this.controller.sourceGeneration === source &&
				this.controller.documentGeneration === document
			)
				this.announce('Updated fill and line formatting.');
		} catch (error) {
			if (request !== this.#request) return;
			this.#pending = false;
			if (!this.#current(context)) {
				this.close();
				return;
			}
			this.#view.apply.disabled = false;
			for (const input of this.#view.fields.values()) input.disabled = false;
			if (!isEditCancellation(error)) this.#view.error.textContent = editErrorMessage(error);
		}
	}
}
