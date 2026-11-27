// The pen strip's settings menu: thickness choices (a radio group), Office's ink colours and the
// standard colours (the shared colour grid), More Colors and Delete Pen. It is a dialog that keeps
// focus: Tab and Shift+Tab move between its three stops (the checked thickness, the colour grid,
// Delete Pen) and wrap; arrow keys move inside a stop.
import { html } from 'lit';
import type { ThemePaletteColumn } from 'ooxml-core/color';
import { INK_PEN_THICKNESSES_MM, inkThicknessPx, type InkPen } from 'ooxml-core/drawingml';
import type { OfficeColorPickEvent } from '../form/color-grid';
import { INK_PEN_COLORS, type InkPensLabels } from './pen-labels';

const choices = (pen: InkPen) =>
	pen.kind === 'highlighter' ? INK_PEN_THICKNESSES_MM.highlighter : INK_PEN_THICKNESSES_MM.pen;

/** The nearest thickness choice of a pen, in millimetres. */
export function inkPenThicknessMm(pen: InkPen): number {
	return choices(pen).reduce((best, mm) =>
		Math.abs(inkThicknessPx(mm) - pen.width) < Math.abs(inkThicknessPx(best) - pen.width)
			? mm
			: best,
	);
}

export interface PenMenuHost {
	labels: InkPensLabels;
	pen: InkPen;
	/** The pen's name, which names the dialog. */
	name: string;
	/** A thickness in the product's language. */
	mm(value: number): string;
	change(change: { color?: string; width?: number }): void;
	/** Delete Pen is off for the last pen. */
	removable: boolean;
	remove(): void;
}

/** Office's ink colours, as the colour grid's columns. */
function palette(labels: InkPensLabels): ThemePaletteColumn[] {
	return INK_PEN_COLORS.map(({ value, name }) => {
		const label = labels.colorNames?.[name] ?? name;
		return { name: label, base: { hex: value.toLowerCase(), label }, variants: [] };
	});
}

function moreColors(event: Event): void {
	const input = (event.currentTarget as HTMLElement)
		.closest('.menu')
		?.querySelector<HTMLInputElement>('.more-input');
	if (!input) return;
	try {
		input.showPicker();
	} catch {
		input.click();
	}
}

/** The menu's keys: arrows move the checked thickness; Tab keeps focus inside the dialog. */
function menuKey(event: KeyboardEvent): void {
	const menu = event.currentTarget as HTMLElement;
	if (event.key === 'Tab') {
		const stops = [
			menu.querySelector<HTMLElement>('.thickness[tabindex="0"]'),
			menu.querySelector<HTMLElement>('office-ui-color-grid'),
			menu.querySelector<HTMLButtonElement>('.remove:not(:disabled)'),
		].filter((stop): stop is HTMLElement => Boolean(stop));
		const root = menu.getRootNode() as ShadowRoot;
		const at = stops.findIndex((stop) => stop === root.activeElement);
		const next = stops[(at + (event.shiftKey ? -1 : 1) + stops.length) % stops.length];
		event.preventDefault();
		event.stopPropagation();
		next?.focus();
		return;
	}
	const radios = [...menu.querySelectorAll<HTMLButtonElement>('.thickness')];
	const at = radios.indexOf(event.target as HTMLButtonElement);
	const step = { ArrowRight: 1, ArrowDown: 1, ArrowLeft: -1, ArrowUp: -1 }[event.key];
	if (at < 0 || !step) return;
	event.preventDefault();
	event.stopPropagation();
	const next = radios[(at + step + radios.length) % radios.length];
	next?.click();
	menu.querySelector<HTMLButtonElement>(`.thickness[data-mm="${next?.dataset.mm}"]`)?.focus();
}

export function penMenu(host: PenMenuHost) {
	const { labels, pen } = host;
	const current = inkPenThicknessMm(pen);
	return html`<div class="menu" role="dialog" aria-label=${host.name} @keydown=${menuKey}
		><div class="heading" id="thickness">${labels.thickness}</div
		><div class="thicknesses" role="radiogroup" aria-labelledby="thickness"
			>${choices(pen).map((mm) => {
				const px = inkThicknessPx(mm);
				const on = current === mm;
				const name = `${host.mm(mm)} mm`;
				return html`<button
					type="button"
					class="thickness"
					role="radio"
					data-mm=${mm}
					tabindex=${on ? '0' : '-1'}
					aria-checked=${String(on)}
					aria-label=${name}
					title=${name}
					@click=${() => host.change({ width: px })}
					><span style=${`height:${Math.max(1, Math.min(14, px))}px;color:${pen.color}`}></span
				></button>`;
			})}</div
		><office-ui-color-grid
			class="colors"
			.palette=${palette(labels)}
			.value=${pen.color.toLowerCase()}
			label=${labels.colors}
			theme-heading=${labels.colors}
			standard-heading=${labels.standardColors ?? 'Standard Colors'}
			more-label=${labels.moreColors}
			@office-color-pick=${(event: OfficeColorPickEvent) =>
				host.change({ color: event.detail.color.toUpperCase() })}
			@office-color-more=${moreColors}
		></office-ui-color-grid
		><input
			class="more-input"
			type="color"
			tabindex="-1"
			aria-hidden="true"
			.value=${pen.color.toLowerCase()}
			@change=${(event: Event) =>
				host.change({ color: (event.target as HTMLInputElement).value.toUpperCase() })}
		/><button
			type="button"
			class="remove"
			?disabled=${!host.removable}
			@click=${() => host.remove()}
			>${labels.remove}</button
		></div
	>`;
}
