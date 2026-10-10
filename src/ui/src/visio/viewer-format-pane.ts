import type { VisioShapeFormatEdit } from 'ooxml-core/visio';
import { visioShapeFormattingState } from 'ooxml-core/visio/ui';
import type { OfficeColorPick } from '../controls';
import type { ViewerController, ViewerState } from './controller';
import {
	DASH_TYPES,
	createFormatPaneView,
	type FormatPaneSection,
	type PaneNumber,
	type PaneTarget,
} from './format-pane-view';
import type { VisioFormattingAction, VisioRibbonAction } from './ribbon-action';
import { colorAction, commonColor, pickedThemeColor } from './ribbon-color-menu';
import { styleSelection, type ViewerColorMenus } from './viewer-color-menus';

type Patch = Omit<VisioShapeFormatEdit, 'type' | 'pageId' | 'shapeId'>;
const NUMBERS: Readonly<Record<PaneNumber, { label: string; max: number }>> = {
	fillTransparency: { label: 'a fill transparency', max: 100 },
	lineTransparency: { label: 'a line transparency', max: 100 },
	lineWeight: { label: 'a line width', max: 150 },
};
const round = (value: number) => String(Math.round(value * 100) / 100);

function reasonFor(state: ViewerState): string {
	if (state.loading) return 'Opening diagram.';
	if (!state.document) return 'Open a drawing to format its shapes.';
	if (!state.edit.sourceAvailable) return 'Open a .vsdx file to format shapes.';
	if (!state.selectedShapes.length) return 'Select a shape to format.';
	return styleSelection(state)
		? ''
		: 'Formatting requires a local shape without a master, group, foreign image, or layer membership.';
}

/**
 * Visio's Format Shape task pane (Home > Shape Styles launcher): Fill (none or solid, colour,
 * transparency) and Line (none or solid, colour, transparency, width, dash type). It follows the
 * selection and every change is applied at once as one undoable `format-shape` edit per shape,
 * through the same action the ribbon uses. Effects and fill patterns keep their dialogs, linked
 * from the foot of the pane.
 */
export class ViewerFormatPane {
	readonly #view;
	/** The colour to return to when a fill or line is switched back from none. */
	#last: Record<PaneTarget, string> = { fill: '#ffffff', line: '#000000' };
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly colors: ViewerColorMenus,
		private readonly run: (action: VisioFormattingAction) => void,
		private readonly raise: (action: VisioRibbonAction) => void,
		private readonly reveal: () => void,
		private readonly announce: (message: string) => void,
	) {
		this.#view = createFormatPaneView(root.ownerDocument);
		for (const { value, label } of DASH_TYPES) {
			const option = root.ownerDocument.createElement('option');
			option.value = value;
			option.textContent = label;
			this.#view.dash.append(option);
		}
		// A harness without the workspace chrome still gets the view.
		(root.querySelector('.inspector-pane') ?? root).append(this.#view.section);
	}
	/** Show the pane; it explains itself when nothing can be formatted. */
	show(): void {
		this.render(this.controller.state);
		this.reveal();
	}
	#format(patch: Patch): void {
		this.run({ type: 'shape-format', patch });
	}
	wire(): () => void {
		const Abort = this.root.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		const view = this.#view;
		for (const target of ['fill', 'line'] as const) {
			const part = view[target];
			part.none.addEventListener(
				'change',
				() => this.#format(target === 'fill' ? { fillColor: 'none' } : { linePattern: 0 }),
				options,
			);
			part.solid.addEventListener(
				'change',
				() => this.#format(target === 'fill' ? { fillColor: this.#last.fill } : { linePattern: 1 }),
				options,
			);
			part.color.addEventListener(
				'click',
				() => this.#toggleGrid(part, part.grid.hasAttribute('hidden')),
				options,
			);
			part.grid.addEventListener(
				'office-color-pick',
				(event) => {
					event.stopPropagation();
					this.#toggleGrid(part, false);
					const pick = (event as CustomEvent<OfficeColorPick>).detail;
					this.run(colorAction(target, pick.color, pickedThemeColor(pick)));
				},
				options,
			);
			part.grid.addEventListener(
				'office-color-more',
				(event) => {
					event.stopPropagation();
					this.#toggleGrid(part, false);
					this.colors.moreColors.open(part.grid.value ?? undefined, (color) => {
						this.colors.remember(color);
						this.run(colorAction(target, color));
					});
				},
				options,
			);
		}
		for (const field of Object.keys(NUMBERS) as PaneNumber[])
			view.numbers[field].addEventListener('change', () => this.#number(field), options);
		view.dash.addEventListener(
			'change',
			() => view.dash.value && this.#format({ linePattern: Number(view.dash.value) }),
			options,
		);
		view.effects.addEventListener(
			'click',
			() => this.raise({ type: 'format-shape-pane' }),
			options,
		);
		view.patterns.addEventListener(
			'click',
			() => this.raise({ type: 'paint-properties' }),
			options,
		);
		// Typing in the pane is not a canvas shortcut.
		view.section.addEventListener(
			'keydown',
			(event) => {
				if (event.key !== 'Escape' && event.key !== 'Tab') event.stopPropagation();
			},
			options,
		);
		// The Shape Styles group's dialog launcher.
		this.root.addEventListener(
			'office-command',
			(event) => {
				if ((event as CustomEvent<{ command?: unknown }>).detail?.command === 'shape-styles-dialog')
					this.show();
			},
			options,
		);
		return () => {
			events.abort();
			view.section.remove();
		};
	}
	#toggleGrid(part: FormatPaneSection, open: boolean): void {
		part.grid.hidden = !open;
		part.color.setAttribute('aria-expanded', String(open));
		if (open) part.grid.focus();
	}
	#number(field: PaneNumber): void {
		const input = this.#view.numbers[field];
		const { label, max } = NUMBERS[field];
		const value = Number(input.value);
		if (input.value.trim() === '' || !Number.isFinite(value) || value < 0 || value > max) {
			this.announce(`Enter ${label} from 0 to ${max}.`);
			this.render(this.controller.state);
			return;
		}
		this.#format({ [field]: value } as Patch);
	}
	render(state: ViewerState): void {
		const view = this.#view;
		const reason = reasonFor(state);
		const selection = reason ? undefined : styleSelection(state);
		const shapes = selection?.shapes ?? [];
		const common = visioShapeFormattingState(shapes);
		const disabled = !!reason || state.edit.busy;
		view.hint.textContent = reason;
		view.hint.hidden = !reason;
		const active = this.root.activeElement;
		const put = (input: HTMLInputElement | HTMLSelectElement, value: string) => {
			// Do not replace what is being typed.
			if (input !== active || disabled) input.value = value;
			input.disabled = disabled;
		};
		const fill = commonColor(shapes.map((shape) => shape.style.fill));
		const noLine = shapes.length > 0 && common.linePattern === 0;
		const line = commonColor(shapes.map((shape) => shape.style.lineColor));
		const current = {
			fill: { none: fill === 'none', color: fill === 'none' ? undefined : fill, known: !!fill },
			line: { none: noLine, color: line, known: common.linePattern !== undefined },
		};
		for (const target of ['fill', 'line'] as const) {
			const part = view[target];
			const { none, color, known } = current[target];
			if (color && /^#[0-9a-f]{6}$/.test(color)) this.#last[target] = color;
			// While an edit runs the model still has the old value: leave the choice just made alone.
			if (!state.edit.busy) {
				part.none.checked = known && none;
				part.solid.checked = known && !none;
				part.details.hidden = shapes.length > 0 && none;
			}
			part.none.disabled = part.solid.disabled = disabled;
			part.color.disabled = disabled;
			part.color.title = `${target === 'fill' ? 'Fill' : 'Line'} color${color ? `: ${color.toUpperCase()}` : ''}`;
			(part.color.firstElementChild as HTMLElement).style.background = color ?? 'transparent';
			part.color.toggleAttribute('data-mixed', !color);
			if (disabled || none) this.#toggleGrid(part, false);
		}
		put(
			view.numbers.fillTransparency,
			common.fillTransparency === undefined ? '' : String(common.fillTransparency),
		);
		put(
			view.numbers.lineTransparency,
			common.lineTransparency === undefined ? '' : String(common.lineTransparency),
		);
		const widths = shapes.map((shape) => round(shape.style.lineWidth * 72));
		put(view.numbers.lineWeight, commonColor(widths) ?? '');
		put(
			view.dash,
			common.linePattern === undefined || common.linePattern === 0
				? ''
				: String(common.linePattern),
		);
		view.effects.disabled = view.patterns.disabled = disabled;
	}
}
