import type { VisioEdit, VisioTextFormatEdit } from 'ooxml-core/visio';
import {
	visioFormattingShape,
	visioStyleFormattingShape,
	visioFontFamilies,
	visioOrderingShape,
	visioTextFormattingState,
	visioTextIndentCommand,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import type { VisioFormattingAction } from './ribbon-action';
import type { RibbonCommand } from './ribbon-parts';
import { paintOptions, fontColorOptions } from './ribbon-style-options';

type Combo = RibbonCommand & {
	value: string;
	options: { value: string; label: string; disabled?: boolean }[];
};
const sizes = [6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
const toggles = ['bold', 'italic', 'underline', 'strikethrough'] as const;
const horizontal = ['left', 'center', 'right'] as const;
const vertical = ['top', 'middle', 'bottom'] as const;

/** Selection-aware ribbon presentation; all writes go through the controller's edit worker. */
export class ViewerFormatting {
	#fontColor = '#000000';
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, success: string) => void,
	) {}

	run(action: VisioFormattingAction): void {
		const state = this.controller.state;
		if (!state.edit.sourceAvailable || state.loading || state.edit.busy) return;
		const page = state.document?.pages[state.pageIndex];
		const selections = state.selectedShapes;
		if (
			!page ||
			!selections.length ||
			selections.some((selection) => selection.pageId && selection.pageId !== page.id)
		)
			return;
		if (action.type === 'shape-order' && selections.length > 1) return;
		const candidates = selections.map(({ id }) =>
			action.type === 'shape-order'
				? visioOrderingShape(page, id)
				: action.type === 'shape-format'
					? visioStyleFormattingShape(page, id)
					: visioFormattingShape(page, id),
		);
		if (candidates.some((shape) => !shape)) return;
		const shapes = candidates.filter((shape) => !!shape);
		const aggregate = visioTextFormattingState(shapes);
		if (action.type === 'font-color' && action.value) this.#fontColor = action.value;
		const commands: VisioEdit[] = shapes.map((shape) => {
			const target = { pageId: page.id, shapeId: shape.id };
			if (action.type === 'shape-order')
				return { type: 'reorder-shape', ...target, order: action.order };
			if (action.type === 'shape-format')
				return { type: 'format-shape', ...target, ...action.patch };
			if (action.type === 'text-indent')
				return visioTextIndentCommand(page, shape.id, action.direction)!;
			const patch: Omit<VisioTextFormatEdit, 'type' | 'pageId' | 'shapeId'> = {};
			switch (action.type) {
				case 'text-toggle':
					patch[action.property] = !aggregate[action.property];
					break;
				case 'text-bullets':
					patch.bullets = !aggregate.bullets;
					break;
				case 'font-color':
					patch.fontColor = this.#fontColor;
					break;
				case 'font-family':
					patch.fontFamily = action.value;
					break;
				case 'font-size':
					patch.fontSize = action.value;
					break;
				case 'font-step': {
					const current = shape.text.fontSize * 72;
					patch.fontSize =
						action.direction === 1
							? (sizes.find((size) => size > current + 0.001) ??
								Math.min(1000, Math.ceil(current * 1.2)))
							: ([...sizes].reverse().find((size) => size < current - 0.001) ??
								Math.max(1, Math.floor(current / 1.2)));
					break;
				}
				case 'text-align':
					if (action.axis === 'horizontal') patch.horizontalAlign = action.value;
					else patch.verticalAlign = action.value;
					break;
			}
			return { type: 'format-text', ...target, ...patch };
		});
		if (commands.some((command) => !command)) return;
		this.edit(
			() => this.controller.applyEdits(commands),
			action.type === 'shape-order' ? 'Updated shape order.' : 'Updated shape formatting.',
		);
	}

	render(state: ViewerState): void {
		const page = state.document?.pages[state.pageIndex];
		const selections = state.selectedShapes;
		const selection = state.selectedShape;
		const currentPage = selections.every((item) => !item.pageId || item.pageId === page?.id);
		const candidates =
			page && currentPage ? selections.map((item) => visioFormattingShape(page, item.id)) : [];
		const styleCandidates =
			page && currentPage ? selections.map((item) => visioStyleFormattingShape(page, item.id)) : [];
		const shapes = candidates.filter((shape) => !!shape);
		const styleShapes = styleCandidates.filter((shape) => !!shape);
		const shape = shapes.length === selections.length ? shapes.at(-1) : undefined;
		const styleShape = styleShapes.length === selections.length ? styleShapes.at(-1) : undefined;
		const aggregate = visioTextFormattingState(shapes);
		const baseReason = state.loading
			? 'Opening diagram.'
			: state.edit.busy
				? 'Updating diagram.'
				: !state.edit.sourceAvailable
					? 'Open a .vsdx file to edit formatting.'
					: !selection
						? 'Select a shape to edit formatting.'
						: '';
		const reason =
			baseReason ||
			(!shape
				? 'Text formatting requires a local shape with uniform text styles, without a master, group, or layer membership.'
				: '');
		const styleReason =
			baseReason ||
			(!styleShape
				? 'Formatting requires a local shape without a master, group, foreign image, or layer membership.'
				: '');
		const set = (element: RibbonCommand | null, disabled: string, pressed?: boolean) => {
			if (!element) return;
			element.disabled = !!disabled;
			const label = element.getAttribute('label') ?? element.getAttribute('aria-label') ?? '';
			element.title = disabled
				? `${label}: ${disabled}`
				: `${label}: source formulas and protection can refuse this edit.`;
			if (pressed !== undefined) element.setAttribute('pressed', String(pressed));
		};
		const button = (id: string) => this.root.querySelector<RibbonCommand>(`[command="${id}"]`);
		for (const property of toggles) set(button(property), reason, !!shape && aggregate[property]);
		for (const value of horizontal)
			set(button(`align-${value}`), reason, !!shape && aggregate.horizontalAlign === value);
		set(button('justify'), reason, !!shape && aggregate.horizontalAlign === 'justify');
		set(button('bullets'), reason, !!shape && aggregate.bullets);
		const indentReason =
			reason ||
			(!page || !shapes.every((item) => visioTextIndentCommand(page, item.id, 'increase'))
				? 'Indent requires uniform supported paragraph indents.'
				: '');
		set(button('indent-increase'), indentReason);
		set(
			button('indent-decrease'),
			indentReason || (!aggregate.canIndentDecrease ? 'The selected text has no left indent.' : ''),
		);
		set(button('font-color'), reason);
		for (const item of fontColorOptions()) {
			const el = button(item.id);
			set(el, reason);
			if (el && item.action?.type === 'font-color')
				el.setAttribute('checked', String(!!shape && aggregate.fontColor === item.action.value));
		}
		for (const value of vertical)
			set(button(`align-${value}`), reason, !!shape && aggregate.verticalAlign === value);
		set(
			button('grow-font'),
			reason ||
				(shape && shapes.every((item) => item.text.fontSize * 72 >= 1000)
					? 'Maximum supported size is 1000 pt.'
					: ''),
		);
		set(
			button('shrink-font'),
			reason ||
				(shape && shapes.every((item) => item.text.fontSize * 72 <= 1)
					? 'Minimum supported size is 1 pt.'
					: ''),
		);
		const font = this.root.querySelector<Combo>('[data-combo="font"]');
		const size = this.root.querySelector<Combo>('[data-combo="font-size"]');
		const families = state.document ? visioFontFamilies(state.document) : [];
		set(font, reason || (!families.length ? 'The drawing has no supported font definitions.' : ''));
		set(size, reason);
		if (font) {
			font.options = families.map((value) => ({ value, label: value }));
			const current = aggregate.fontFamily ?? '';
			if (!families.includes(current))
				font.options.unshift({
					value: current,
					label: current || (selections.length ? 'Mixed' : 'Font'),
					disabled: true,
				});
			font.value = aggregate.fontFamily ?? '';
		}
		if (size) {
			const current =
				shape && aggregate.fontSize !== undefined
					? Number((aggregate.fontSize * 72).toFixed(2))
					: undefined;
			size.options = [...new Set([...sizes, ...(current === undefined ? [] : [current])])]
				.sort((a, b) => a - b)
				.map((value) => ({ value: String(value), label: `${value} pt` }));
			if (current === undefined)
				size.options.unshift({
					value: '',
					label: selections.length ? 'Mixed' : 'Size',
					disabled: true,
				});
			size.value = current === undefined ? '' : String(current);
		}
		for (const target of ['fill', 'line'] as const) {
			set(this.root.querySelector<RibbonCommand>(`[data-menu="${target}"]`), styleReason);
			for (const item of paintOptions(target)) {
				if (item.items) {
					set(this.root.querySelector<RibbonCommand>(`[data-menu="${item.id}"]`), styleReason);
					for (const weight of item.items) {
						const el = button(weight.id);
						set(el, styleReason);
						if (el && weight.action?.type === 'shape-format') {
							const points = weight.action.patch.lineWeight ?? 0;
							el.setAttribute(
								'checked',
								String(
									!!styleShape &&
										styleShapes.every(
											(itemShape) => Math.abs(itemShape.style.lineWidth * 72 - points) < 0.001,
										),
								),
							);
						}
					}
				} else {
					const el = button(item.id);
					set(el, styleReason);
					if (el && item.action?.type === 'shape-format') {
						const color =
							target === 'fill' ? item.action.patch.fillColor : item.action.patch.lineColor;
						el.setAttribute(
							'checked',
							String(
								!!styleShape &&
									styleShapes.every(
										(itemShape) =>
											(target === 'fill' ? itemShape.style.fill : itemShape.style.lineColor) ===
											color,
									),
							),
						);
					}
				}
			}
		}
		const ordering =
			page && selection && currentPage ? visioOrderingShape(page, selection.id) : undefined;
		const orderReason =
			state.loading || state.edit.busy || !state.edit.sourceAvailable || !selection
				? baseReason
				: selections.length > 1
					? 'Select one shape to change its drawing order.'
					: !ordering
						? 'Ordering requires a local shape without a master, group, or layer.'
						: '';
		const index = ordering && page ? page.shapes.indexOf(ordering) : -1;
		for (const id of ['bring-to-front', 'bring-forward'])
			set(
				button(id),
				orderReason ||
					(index === (page?.shapes.length ?? 0) - 1 ? 'The shape is already at the front.' : ''),
			);
		for (const id of ['send-to-back', 'send-backward'])
			set(button(id), orderReason || (index === 0 ? 'The shape is already at the back.' : ''));
	}
}
