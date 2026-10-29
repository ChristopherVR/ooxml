// State and interaction of legacy form controls. The linked cell is the source of truth: a check
// box shows TRUE/FALSE, an option button group the 1-based number of the selected button, a
// drop-down or list box the 1-based selected item, and a spin button or scroll bar its number.
// The functions here are pure; the editor reads the cell, calls them and writes the result back.
import type { FormControlObject } from '../model';

const DEFAULT_MAX = 30000;

/** The value a control holds, read from its linked cell (`undefined` for an empty or unlinked one). */
export interface ControlState {
	checked: 'checked' | 'mixed' | 'unchecked';
	/** 1-based item (list or drop-down) or option number. */
	selected: number;
	value: number;
}

export type LinkedValue = number | boolean | string | undefined;

/** Controls that change the workbook when used. */
export const isInteractiveControl = (control: FormControlObject): boolean =>
	control.controlType === 'checkbox' ||
	control.controlType === 'radio' ||
	control.controlType === 'dropdown' ||
	control.controlType === 'list' ||
	control.controlType === 'spin' ||
	control.controlType === 'scroll';

export function controlBounds(control: FormControlObject): { min: number; max: number } {
	const min = control.min ?? 0;
	return { min, max: Math.max(min, control.max ?? DEFAULT_MAX) };
}

const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value));

/**
 * The radio group a button belongs to: the buttons from the nearest preceding `firstButton`
 * (inclusive) up to the next one, in drawing order. The group's link is the first button's.
 */
export function radioGroup(
	controls: readonly FormControlObject[],
	control: FormControlObject,
): { buttons: FormControlObject[]; index: number; linkedCell: string | undefined } {
	const at = controls.indexOf(control);
	let start = at;
	while (start > 0 && !(controls[start]?.firstButton ?? false)) {
		const previous = controls[start - 1];
		if (previous?.controlType !== 'radio' && previous?.controlType !== 'groupBox') break;
		start--;
	}
	const buttons: FormControlObject[] = [];
	for (let i = start; i < controls.length; i++) {
		const candidate = controls[i]!;
		if (candidate.controlType === 'groupBox') continue;
		if (candidate.controlType !== 'radio' || (i > start && candidate.firstButton)) break;
		buttons.push(candidate);
	}
	return {
		buttons,
		index: Math.max(0, buttons.indexOf(control)),
		linkedCell: buttons.find((button) => button.linkedCell)?.linkedCell,
	};
}

/** The state of a control given the value of its linked cell. */
export function controlState(
	control: FormControlObject,
	linked: LinkedValue,
	radioIndex = 0,
): ControlState {
	const number = typeof linked === 'number' ? linked : Number.NaN;
	const bounds = controlBounds(control);
	const hasLink = control.linkedCell !== undefined && linked !== undefined;
	let checked: ControlState['checked'] = control.checked ?? 'unchecked';
	if (control.controlType === 'checkbox' && hasLink)
		checked =
			linked === true || linked === 1
				? 'checked'
				: linked === false || linked === 0
					? 'unchecked'
					: 'mixed';
	if (control.controlType === 'radio' && hasLink && Number.isFinite(number))
		checked = number === radioIndex + 1 ? 'checked' : 'unchecked';
	return {
		checked,
		selected: hasLink && Number.isFinite(number) ? Math.trunc(number) : (control.selected ?? 0),
		value: clamp(
			hasLink && Number.isFinite(number) ? number : (control.value ?? bounds.min),
			bounds.min,
			bounds.max,
		),
	};
}

export type FormControlAction =
	| { type: 'toggle' }
	| { type: 'select'; index: number }
	| { type: 'step'; delta: number }
	| { type: 'page'; direction: 1 | -1 }
	| { type: 'set'; value: number };

/**
 * The value to write to the linked cell for an action, or undefined when nothing changes (no
 * link, a disabled control, an unsupported action). `radioIndex` is the button's place in its group.
 */
export function applyControlAction(
	control: FormControlObject,
	current: ControlState,
	action: FormControlAction,
	radioIndex = 0,
): number | boolean | undefined {
	const bounds = controlBounds(control);
	switch (control.controlType) {
		case 'checkbox':
			return action.type === 'toggle' ? current.checked !== 'checked' : undefined;
		case 'radio':
			return action.type === 'toggle' || action.type === 'select' ? radioIndex + 1 : undefined;
		case 'dropdown':
		case 'list':
			return action.type === 'select' && action.index >= 1 ? Math.trunc(action.index) : undefined;
		case 'spin':
		case 'scroll': {
			const step = control.step ?? 1;
			if (action.type === 'step')
				return clamp(current.value + action.delta * step, bounds.min, bounds.max);
			if (action.type === 'page')
				return clamp(
					current.value + action.direction * (control.page ?? 10),
					bounds.min,
					bounds.max,
				);
			if (action.type === 'set') return clamp(Math.round(action.value), bounds.min, bounds.max);
			return undefined;
		}
		default:
			return undefined;
	}
}

/** The value of a scroll bar whose thumb sits at `fraction` (0..1) of its track. */
export function controlValueAtFraction(control: FormControlObject, fraction: number): number {
	const { min, max } = controlBounds(control);
	return clamp(Math.round(min + clamp(fraction, 0, 1) * (max - min)), min, max);
}
