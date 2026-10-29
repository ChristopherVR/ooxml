// Worksheet shapes and form controls over the grid. The core draws both as SVG
// (`shapeSvg`, `formControlSvg`); this module places that markup and turns presses on a form
// control into a write to its linked cell (the cell stays the source of truth, so undo, collaboration
// and recalculation need nothing more).
import {
	applyControlAction,
	controlLinkCell,
	controlListItems,
	controlStateIn,
	controlValueAtFraction,
	formControlSvg,
	shapeSvg,
	type FormControlAction,
	type FormControlObject,
	type ShapeNode,
} from 'ooxml-core/xlsx';
import { h, svgNode } from './dom';
import type { GridView } from './grid-view';
import { trackPointerGesture } from '../../pointer-gesture';

/** Paints a shape, connector or group into `node` from its drawing box. */
export function paintShape(
	view: GridView,
	node: HTMLElement,
	shape: ShapeNode,
	box: { w: number; h: number },
	idPrefix: string,
): void {
	node.classList.add('xg-shape');
	node.setAttribute('role', 'img');
	node.setAttribute('aria-label', shape.name ?? '');
	node.append(
		svgNode(
			view.doc,
			shapeSvg(shape, box.w, box.h, {
				theme: view.workbook()?.theme,
				scale: view.zoom() / 100,
				idPrefix,
			}),
		),
	);
}

/** Paints a form control into `node` from its linked cell, with the drop-down list when open. */
export function paintFormControl(
	view: GridView,
	node: HTMLElement,
	control: FormControlObject,
	box: { w: number; h: number },
	open: boolean,
): void {
	const workbook = view.workbook();
	const sheet = view.sheet();
	node.classList.add('xg-formcontrol');
	node.dataset.controlType = control.controlType;
	if (!workbook || !sheet) return;
	const controls = sheet.drawings.filter((d): d is FormControlObject => d.kind === 'formControl');
	const index = view.sheetIndex();
	const { state } = controlStateIn(workbook, index, controls, control);
	const items = controlListItems(workbook, index, control);
	const markup = formControlSvg(control, state, box.w, box.h, {
		items,
		scale: view.zoom() / 100,
	});
	if (!markup) {
		node.classList.add('xg-obj-unsupported');
		node.textContent = control.name ?? control.controlType;
		return;
	}
	node.append(svgNode(view.doc, markup));
	node.setAttribute('role', 'group');
	node.setAttribute('aria-label', control.caption ?? control.name ?? control.controlType);
	if (open && control.controlType === 'dropdown') {
		const list = h(view.doc, 'div', 'xg-fc-pop');
		list.style.top = `${box.h}px`;
		list.style.width = `${box.w}px`;
		items.forEach((item, i) => {
			const row = h(view.doc, 'div', 'xg-fc-item');
			row.dataset.controlPart = 'item';
			row.dataset.controlIndex = String(i + 1);
			row.textContent = item;
			if (state.selected === i + 1) row.classList.add('xg-fc-on');
			list.append(row);
		});
		node.append(list);
	}
}

/**
 * Handles a press on a form control. Returns true when the press hit a part that reacts (so the
 * caller neither selects nor drags the object), false to let the press select and move it.
 */
export function pressFormControl(
	view: GridView,
	event: PointerEvent,
	control: FormControlObject,
	node: HTMLElement,
	hooks: { toggleOpen: (open: boolean) => void; isOpen: () => boolean; refresh: () => void },
): boolean {
	const part = (event.target as Element).closest<HTMLElement | SVGElement>('[data-control-part]');
	const name = part?.dataset.controlPart;
	if (!part || !name) {
		hooks.toggleOpen(false);
		return false;
	}
	const workbook = view.workbook();
	const sheet = view.sheet();
	const session = view.ctx.session();
	if (!workbook || !sheet || !session || view.ctx.readOnly()) return true;
	const controls = sheet.drawings.filter((d): d is FormControlObject => d.kind === 'formControl');
	const sheetIndex = view.sheetIndex();
	const resolved = controlStateIn(workbook, sheetIndex, controls, control);
	const target = controlLinkCell(workbook, sheetIndex, resolved.link);
	const write = (action: FormControlAction): void => {
		const value = applyControlAction(
			control,
			controlStateIn(workbook, sheetIndex, controls, control).state,
			action,
			resolved.radioIndex,
		);
		if (value === undefined || !target) return;
		try {
			session.setCellValue(target.sheet, target.row, target.col, value);
		} catch (error) {
			view.ctx.toast(view.ctx.t(error instanceof Error ? error.message : String(error)), 'warning');
		}
	};
	if (name === 'drop') {
		hooks.toggleOpen(!hooks.isOpen());
		return true;
	}
	hooks.toggleOpen(false);
	if (name === 'toggle') write({ type: 'toggle' });
	else if (name === 'item') write({ type: 'select', index: Number(part.dataset.controlIndex) });
	else if (name === 'step-up' || name === 'step-down')
		write({ type: 'step', delta: name === 'step-up' ? 1 : -1 });
	else if (name === 'track' || name === 'thumb') {
		const svg = node.querySelector('svg');
		const track = node.querySelector<SVGGraphicsElement>('[data-control-part="track"]');
		const thumb = node.querySelector<SVGGraphicsElement>('[data-control-part="thumb"]');
		if (!svg || !track || !thumb) return true;
		const horizontal = control.horizontal ?? node.clientWidth >= node.clientHeight;
		const rect = svg.getBoundingClientRect();
		const zoom = view.zoom() / 100;
		const trackBox = track.getBBox();
		const thumbBox = thumb.getBBox();
		const span =
			(horizontal ? trackBox.width - thumbBox.width : trackBox.height - thumbBox.height) * zoom ||
			1;
		const coordinate = (e: PointerEvent) =>
			horizontal ? e.clientX - rect.left : e.clientY - rect.top;
		if (name === 'track') {
			const thumbStart = (horizontal ? thumbBox.x : thumbBox.y) * zoom;
			write({ type: 'page', direction: coordinate(event) > thumbStart ? 1 : -1 });
		} else {
			const grabOffset = coordinate(event) - (horizontal ? thumbBox.x : thumbBox.y) * zoom;
			const origin = (horizontal ? trackBox.x : trackBox.y) * zoom;
			trackPointerGesture(node, event.pointerId, {
				move: (e) => {
					const fraction = (coordinate(e) - grabOffset - origin) / span;
					write({ type: 'set', value: controlValueAtFraction(control, fraction) });
					hooks.refresh();
				},
				commit: () => undefined,
				cancel: () => undefined,
			});
		}
	}
	return true;
}
