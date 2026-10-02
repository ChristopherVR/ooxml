// Small one-field dialogs: confirm, text and number prompts, and list pickers. Larger dialogs
// build on the same frame.
import type { EditorContext } from '../context.js';
import { field, invalid, listBox, numberInput, text, textArea, textInput } from './fields.js';
import { showDialog } from './frame.js';

export interface ConfirmProps {
	heading: string;
	message: string;
	okLabel?: string;
	vars?: Record<string, string | number>;
}

/** Resolves true when confirmed. */
export function confirmDialog(
	ctx: EditorContext,
	props: ConfirmProps,
): Promise<boolean | undefined> {
	return showDialog<boolean>(ctx, {
		name: 'confirm',
		heading: props.heading,
		body: text(ctx, props.message, 'xve-message', props.vars),
		okLabel: props.okLabel ?? 'OK',
		submit: () => true,
	});
}

export interface TextPromptProps {
	name: string;
	heading: string;
	label: string;
	value: string;
	note?: string;
	multiline?: boolean;
	/** Returns an English error message, or undefined when the text is acceptable. */
	validate?(value: string): string | undefined;
	apply(value: string): void;
}

export function textPrompt(
	ctx: EditorContext,
	props: TextPromptProps,
): Promise<string | undefined> {
	const input = props.multiline ? textArea(ctx, props.value, 4) : textInput(ctx, props.value);
	const body: HTMLElement[] = [field(ctx, props.label, input)];
	if (props.note) body.push(text(ctx, props.note));
	return showDialog<string>(ctx, {
		name: props.name,
		heading: props.heading,
		body,
		opened: () => {
			input.focus();
			input.select();
		},
		submit: () => {
			const value = input.value;
			const problem = props.validate?.(value);
			if (problem) return invalid(ctx, input, problem);
			props.apply(value);
			return value;
		},
	});
}

export interface NumberPromptProps {
	name: string;
	heading: string;
	label: string;
	value: number;
	min: number;
	max: number;
	step?: number;
	apply(value: number): void;
}

export function numberPrompt(
	ctx: EditorContext,
	props: NumberPromptProps,
): Promise<number | undefined> {
	const input = numberInput(ctx, props.value, props.min, props.max, props.step ?? 1);
	input.type = 'text';
	input.inputMode = 'decimal';
	return showDialog<number>(ctx, {
		name: props.name,
		heading: props.heading,
		body: field(ctx, props.label, input),
		opened: () => {
			input.focus();
			input.select();
		},
		submit: () => {
			const value = Number(input.value.trim().replace(',', '.'));
			if (!input.value.trim() || !Number.isFinite(value) || value < props.min || value > props.max)
				return invalid(ctx, input, 'Enter a number between {min} and {max}.', {
					min: props.min,
					max: props.max,
				});
			props.apply(value);
			return value;
		},
	});
}

export interface PickProps {
	name: string;
	heading: string;
	label: string;
	items: ReadonlyArray<readonly [value: string, label: string]>;
	okLabel?: string;
	apply(value: string): void;
}

/** A list picker (Unhide sheet, Paste Name). */
export function pickDialog(ctx: EditorContext, props: PickProps): Promise<string | undefined> {
	let confirm: (() => void) | undefined;
	const list = listBox(ctx, props.label, undefined, () => confirm?.());
	list.setItems(props.items);
	const first = props.items[0];
	if (first) list.select(first[0]);
	return showDialog<string>(
		ctx,
		{
			name: props.name,
			heading: props.heading,
			body: [text(ctx, props.label, 'xve-field-label'), list.element],
			okLabel: props.okLabel ?? 'OK',
			opened: () => list.element.focus(),
			submit: () => {
				const value = list.value();
				if (value === undefined) return undefined;
				props.apply(value);
				return value;
			},
		},
		(open) => {
			confirm = () => open.element.querySelector<HTMLButtonElement>('.xve-btn-primary')?.click();
		},
	);
}
