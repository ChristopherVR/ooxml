import { html, nothing } from 'lit';
import type { PptxElement } from 'ooxml-core/pptx';
import { insertDialogText } from '../../dialog/insert-dialog-frame';
import { canInteractWithElement } from '../render/element-locks';
import type { RibbonGroupId } from '../render/customization/ribbon-control-ids';
import type { PictureRibbonState } from './picture-ribbon-state';
import type { PptxUiSelectElement } from './select-value';

export interface ContextualControls {
	host: HTMLElement;
	group: RibbonGroupId;
	state: PictureRibbonState;
	nodeIds: readonly string[];
	begin: () => AbortSignal;
	close: () => void;
}

export function ribbonFields(context: ContextualControls) {
	const { group, state } = context;
	const element = state.element;
	const available = Boolean(state.editable && element && canInteractWithElement(element, 'select'));
	const text = (key: string, fallback: string) => insertDialogText(state.translate, key, fallback);
	const unavailable = text('pptx.mediaRibbon.unavailable', 'Select an editable object.');
	const patch = (value: Partial<PptxElement>) => {
		if (available && element) state.patch(element.id, value);
	};
	const id = (name: string) => `${group}.${name}`;
	const description = (reason?: string) => reason ?? (available ? nothing : unavailable);
	const command = (name: string, label: string, run: () => void, reason?: string) =>
		html`<pptx-ui-ribbon-command
			data-ribbon-control=${id(name)}
			label=${label}
			title=${reason ?? (available ? label : unavailable)}
			aria-description=${description(reason)}
			?disabled=${!available || Boolean(reason)}
			compact
			@command-request=${() => {
				if (available && !reason) run();
			}}
		></pptx-ui-ribbon-command>`;
	const number = (
		name: string,
		label: string,
		value: number,
		min: number,
		max: number,
		step: number | 'any',
		write: (number: number) => void,
		reason?: string,
	) =>
		html`<label
			data-ribbon-control=${id(name)}
			title=${reason ?? (available ? label : unavailable)}
			aria-description=${description(reason)}
			>${label}
			<input
				aria-label=${label}
				aria-description=${description(reason)}
				type="number"
				min=${min}
				max=${max}
				step=${step}
				.value=${String(Number.isFinite(value) ? Number(value.toFixed(3)) : 0)}
				?disabled=${!available || Boolean(reason)}
				@change=${(event: Event) => {
					const input = event.currentTarget as HTMLInputElement;
					if (available && !reason && input.validity.valid && Number.isFinite(input.valueAsNumber))
						write(input.valueAsNumber);
				}}
			/>
		</label>`;
	const toggle = (
		name: string,
		label: string,
		value: boolean,
		write: (checked: boolean) => void,
		reason?: string,
	) =>
		html`<pptx-ui-ribbon-toggle
			data-ribbon-control=${id(name)}
			label=${label}
			title=${reason ?? (available ? label : unavailable)}
			aria-description=${description(reason)}
			.checked=${value}
			?disabled=${!available || Boolean(reason)}
			@toggle-request=${({ detail }: CustomEvent<{ checked: boolean }>) => {
				if (available && !reason) write(detail.checked);
			}}
		></pptx-ui-ribbon-toggle>`;
	const select = (
		name: string,
		label: string,
		value: string,
		options: readonly (readonly [string, string])[],
		write: (selected: string) => void,
		reason?: string,
	) =>
		html`<label
			data-ribbon-control=${id(name)}
			title=${reason ?? (available ? label : unavailable)}
			aria-description=${description(reason)}
			>${label}
			<pptx-ui-select
				aria-label=${label}
				aria-description=${description(reason)}
				.value=${value}
				.disabled=${!available || Boolean(reason)}
				@change=${(event: Event) => {
					if (available && !reason) write((event.currentTarget as PptxUiSelectElement).value);
				}}
			>
				${options.map(([key, option]) => html`<option value=${key}>${option}</option>`)}
			</pptx-ui-select></label
		>`;
	const color = (
		name: string,
		label: string,
		value: string,
		write: (hex: string) => void,
		reason?: string,
	) =>
		html`<label
			class="contextual-color"
			data-ribbon-control=${id(name)}
			title=${reason ?? (available ? label : unavailable)}
			aria-description=${description(reason)}
			><input
				type="color"
				aria-label=${label}
				aria-description=${description(reason)}
				.value=${/^#[0-9a-f]{6}$/iu.test(value) ? value : '#000000'}
				?disabled=${!available || Boolean(reason)}
				@change=${(event: Event) => {
					if (available && !reason) write((event.currentTarget as HTMLInputElement).value);
				}}
			/>${label}
		</label>`;
	return { text, unavailable, patch, command, number, toggle, select, color, available };
}

export type ContextualFields = ReturnType<typeof ribbonFields>;
