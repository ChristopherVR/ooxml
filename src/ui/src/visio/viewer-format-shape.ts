import type { VisioShape, VisioShapeFormatEdit } from 'ooxml-core/visio';
import {
	editErrorMessage,
	isEditCancellation,
	visioSelectionIsOnPage,
	visioShapeEffectValues,
	visioShapeFormattingState,
	visioFormatTargetShape,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { createFormatShapeDialog, type FormatShapeField } from './format-shape-dialog';

type Patch = Omit<VisioShapeFormatEdit, 'type' | 'pageId' | 'shapeId'>;
const common = (values: readonly (string | undefined)[]) =>
	values.length && values.every((value) => value === values[0]) ? (values[0] ?? '') : '';
const round = (value: number) => String(Math.round(value * 100) / 100);

/** The selection's current values as field text; mixed values are empty. */
function fieldValues(shapes: readonly VisioShape[]): Map<FormatShapeField, string> {
	const state = visioShapeFormattingState(shapes);
	const effects = shapes.map(visioShapeEffectValues);
	const each = (read: (shape: VisioShape, index: number) => string | undefined) =>
		common(shapes.map(read));
	const hex = (value: string) => (/^#[0-9a-f]{6}$/i.test(value) ? value.toLowerCase() : undefined);
	return new Map<FormatShapeField, string>([
		['fillColor', each((shape) => (shape.style.fill === 'none' ? 'none' : hex(shape.style.fill)))],
		[
			'fillTransparency',
			state.fillTransparency === undefined ? '' : String(state.fillTransparency),
		],
		['lineColor', each((shape) => hex(shape.style.lineColor))],
		['lineWeight', each((shape) => round(shape.style.lineWidth * 72))],
		['linePattern', state.linePattern === undefined ? '' : String(state.linePattern)],
		[
			'lineTransparency',
			state.lineTransparency === undefined ? '' : String(state.lineTransparency),
		],
		['shadow', state.shadowPreset ?? ''],
		['glowSize', each((_, index) => String(effects[index]!.glow.size))],
		['glowColor', each((_, index) => effects[index]!.glow.color)],
		['glowTransparency', each((_, index) => String(effects[index]!.glow.transparency))],
		['softEdges', each((_, index) => String(effects[index]!.softEdges))],
		['reflectionSize', each((_, index) => String(effects[index]!.reflection.size))],
		['reflectionTransparency', each((_, index) => String(effects[index]!.reflection.transparency))],
		['reflectionDistance', each((_, index) => String(effects[index]!.reflection.distance))],
		['reflectionBlur', each((_, index) => String(effects[index]!.reflection.blur))],
	]);
}

/**
 * The Format Shape dialog with Fill, Line and Effects fields, opened by the Effects menu's Options
 * commands and from the Format Shape task pane (which the Shape Styles launcher opens).
 * Every change becomes one undoable format-shape edit per selected shape.
 */
export class ViewerFormatShape {
	readonly #view;
	#initial = new Map<FormatShapeField, string>();
	#state: ViewerState | undefined;
	#active = false;
	#applying = false;
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly announce: (message: string) => void,
	) {
		this.#view = createFormatShapeDialog(root.ownerDocument);
		root.append(this.#view.dialog);
	}
	wire(): () => void {
		this.#active = true;
		const Abort = this.root.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.#view.apply.addEventListener('office-command', () => void this.#apply(), options);
		this.#view.cancel.addEventListener('office-command', () => this.close(), options);
		this.#view.dialog.addEventListener(
			'office-dialog-close',
			() => (this.#state = undefined),
			options,
		);
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
	show(): void {
		const state = this.controller.state;
		const selection = this.#active && !state.edit.busy ? this.#selection(state) : undefined;
		if (!selection) return;
		this.#state = state;
		this.#initial = fieldValues(selection.shapes);
		for (const [field, input] of this.#view.fields) {
			input.value = this.#initial.get(field) ?? '';
			input.disabled = false;
		}
		this.#view.error.textContent = '';
		this.#view.apply.disabled = false;
		this.#view.dialog.show();
	}
	close(): void {
		this.#state = undefined;
		this.#view.dialog.close();
	}
	render(state: ViewerState): void {
		if (
			this.#state &&
			!this.#applying &&
			state.selectedShapes !== this.#state.selectedShapes &&
			!state.edit.busy
		)
			this.close();
	}
	#patch(): Patch {
		const value = (field: FormatShapeField) => this.#view.fields.get(field)!.value.trim();
		const changed = (...fields: FormatShapeField[]) =>
			fields.some((field) => value(field) !== this.#initial.get(field));
		const number = (field: FormatShapeField, label: string, maximum = 100) => {
			const result = Number(value(field));
			if (!value(field) || !Number.isFinite(result) || result < 0 || result > maximum)
				throw new Error(`Enter ${label} from 0 to ${maximum}.`);
			return result;
		};
		const color = (field: FormatShapeField, label: string, none = false) => {
			const result = value(field).toLowerCase();
			if (!(none && result === 'none') && !/^#[0-9a-f]{6}$/.test(result))
				throw new Error(`Enter the ${label} as #RRGGBB${none ? ' or none' : ''}.`);
			return result;
		};
		const patch: Patch = {};
		if (changed('fillColor')) patch.fillColor = color('fillColor', 'fill colour', true);
		if (changed('fillTransparency'))
			patch.fillTransparency = number('fillTransparency', 'a fill transparency');
		if (changed('lineColor')) patch.lineColor = color('lineColor', 'line colour');
		if (changed('lineWeight')) patch.lineWeight = number('lineWeight', 'a line width');
		if (changed('linePattern')) patch.linePattern = number('linePattern', 'a dash type', 23);
		if (changed('lineTransparency'))
			patch.lineTransparency = number('lineTransparency', 'a line transparency');
		if (changed('shadow')) patch.shadow = value('shadow') as NonNullable<Patch['shadow']>;
		if (changed('glowSize', 'glowColor', 'glowTransparency'))
			patch.glow = {
				size: number('glowSize', 'a glow size', 150),
				color: value('glowColor') ? color('glowColor', 'glow colour') : '#000000',
				transparency: number('glowTransparency', 'a glow transparency'),
			};
		if (changed('softEdges')) patch.softEdges = number('softEdges', 'a soft edge size');
		if (changed('reflectionSize', 'reflectionTransparency', 'reflectionDistance', 'reflectionBlur'))
			patch.reflection = {
				size: number('reflectionSize', 'a reflection size'),
				transparency: number('reflectionTransparency', 'a reflection transparency'),
				distance: number('reflectionDistance', 'a reflection distance'),
				blur: number('reflectionBlur', 'a reflection blur'),
			};
		return patch;
	}
	async #apply(): Promise<void> {
		const context = this.#state;
		if (!context || this.controller.state.edit.busy) return;
		try {
			const patch = this.#patch();
			const selection = this.#selection(this.controller.state);
			if (!Object.keys(patch).length || !selection) {
				this.close();
				return;
			}
			this.#view.apply.disabled = true;
			this.#applying = true;
			await this.controller.applySelectionEdits(
				selection.shapes.map((shape) => ({
					type: 'format-shape' as const,
					pageId: selection.page.id,
					shapeId: shape.id,
					...patch,
				})),
			);
			if (this.#state !== context) return;
			this.close();
			this.announce('Updated shape formatting.');
		} catch (error) {
			if (this.#state !== context) return;
			this.#view.apply.disabled = false;
			if (!isEditCancellation(error)) this.#view.error.textContent = editErrorMessage(error);
		} finally {
			this.#applying = false;
		}
	}
}
