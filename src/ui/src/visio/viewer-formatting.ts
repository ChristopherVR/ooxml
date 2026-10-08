import type { VisioEdit, VisioTextFormatEdit } from 'ooxml-core/visio';
import {
	visioFormattingShape,
	visioStyleFormattingShape,
	visioFontFamilies,
	visioOrderingShape,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import type { VisioFormattingAction } from './ribbon-action';
import type { RibbonCommand } from './ribbon-parts';
import { paintOptions } from './ribbon-style-options';

type Combo = RibbonCommand & {
	value: string;
	options: { value: string; label: string; disabled?: boolean }[];
};
const sizes = [6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
const toggles = ['bold', 'italic', 'underline'] as const;
const horizontal = ['left', 'center', 'right'] as const;
const vertical = ['top', 'middle', 'bottom'] as const;

/** Selection-aware ribbon presentation; all writes go through the controller's edit worker. */
export class ViewerFormatting {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, success: string) => void,
	) {}

	run(action: VisioFormattingAction): void {
		const state = this.controller.state;
		if (!state.edit.sourceAvailable || state.loading || state.edit.busy) return;
		const page = state.document?.pages[state.pageIndex];
		const id = state.selectedShape?.id;
		if (!page || !id || (state.selectedShape?.pageId && state.selectedShape.pageId !== page.id))
			return;
		const shape =
			action.type === 'shape-order'
				? visioOrderingShape(page, id)
				: action.type === 'shape-format'
					? visioStyleFormattingShape(page, id)
					: visioFormattingShape(page, id);
		if (!shape) return;
		const target = { pageId: page.id, shapeId: id };
		let command: VisioEdit;
		if (action.type === 'shape-order')
			command = { type: 'reorder-shape', ...target, order: action.order };
		else if (action.type === 'shape-format')
			command = { type: 'format-shape', ...target, ...action.patch };
		else {
			const patch: Omit<VisioTextFormatEdit, 'type' | 'pageId' | 'shapeId'> = {};
			switch (action.type) {
				case 'text-toggle':
					patch[action.property] = !shape.text.runs[0]?.[action.property];
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
			command = { type: 'format-text', ...target, ...patch };
		}
		this.edit(
			() => this.controller.applyEdits([command]),
			action.type === 'shape-order' ? 'Updated shape order.' : 'Updated shape formatting.',
		);
	}

	render(state: ViewerState): void {
		const page = state.document?.pages[state.pageIndex];
		const selection = state.selectedShape;
		const currentPage = !selection?.pageId || selection.pageId === page?.id;
		const shape =
			page && selection && currentPage ? visioFormattingShape(page, selection.id) : undefined;
		const styleShape =
			page && selection && currentPage ? visioStyleFormattingShape(page, selection.id) : undefined;
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
		for (const property of toggles)
			set(button(property), reason, !!shape?.text.runs[0]?.[property]);
		for (const value of horizontal)
			set(button(`align-${value}`), reason, shape?.text.horizontalAlign === value);
		for (const value of vertical)
			set(button(`align-${value}`), reason, shape?.text.verticalAlign === value);
		set(
			button('grow-font'),
			reason ||
				(shape && shape.text.fontSize * 72 >= 1000 ? 'Maximum supported size is 1000 pt.' : ''),
		);
		set(
			button('shrink-font'),
			reason || (shape && shape.text.fontSize * 72 <= 1 ? 'Minimum supported size is 1 pt.' : ''),
		);
		const font = this.root.querySelector<Combo>('[data-combo="font"]');
		const size = this.root.querySelector<Combo>('[data-combo="font-size"]');
		const families = state.document ? visioFontFamilies(state.document) : [];
		set(font, reason || (!families.length ? 'The drawing has no supported font definitions.' : ''));
		set(size, reason);
		if (font) {
			font.options = families.map((value) => ({ value, label: value }));
			const current = shape?.text.fontFamily ?? '';
			if (!families.includes(current))
				font.options.unshift({ value: current, label: current || 'Font', disabled: true });
			font.value = shape?.text.fontFamily ?? '';
		}
		if (size) {
			const current = shape ? Number((shape.text.fontSize * 72).toFixed(2)) : undefined;
			size.options = [...new Set([...sizes, ...(current === undefined ? [] : [current])])]
				.sort((a, b) => a - b)
				.map((value) => ({ value: String(value), label: `${value} pt` }));
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
						if (el && weight.action?.type === 'shape-format')
							el.setAttribute(
								'checked',
								String(
									Math.abs(
										(styleShape?.style.lineWidth ?? 0) * 72 - (weight.action.patch.lineWeight ?? 0),
									) < 0.001,
								),
							);
					}
				} else {
					const el = button(item.id);
					set(el, styleReason);
					if (el && item.action?.type === 'shape-format')
						el.setAttribute(
							'checked',
							String(
								target === 'fill'
									? styleShape?.style.fill === item.action.patch.fillColor
									: styleShape?.style.lineColor === item.action.patch.lineColor,
							),
						);
				}
			}
		}
		const ordering =
			page && selection && currentPage ? visioOrderingShape(page, selection.id) : undefined;
		const orderReason =
			state.loading || state.edit.busy || !state.edit.sourceAvailable || !selection
				? baseReason
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
