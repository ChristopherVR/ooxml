import type { OfficeUiColorGrid } from '../controls';
import { colorGrid } from './ribbon-color-menu';

export type PaneTarget = 'fill' | 'line';
export type PaneNumber = 'fillTransparency' | 'lineTransparency' | 'lineWeight';

export interface FormatPaneSection {
	set: HTMLFieldSetElement;
	none: HTMLInputElement;
	solid: HTMLInputElement;
	/** Shows the current colour and opens the grid under it. */
	color: HTMLButtonElement;
	grid: OfficeUiColorGrid;
	/** Everything that only applies to a solid fill or line. */
	details: HTMLElement;
}
export interface FormatPaneView {
	section: HTMLElement;
	hint: HTMLElement;
	fill: FormatPaneSection;
	line: FormatPaneSection;
	numbers: Record<PaneNumber, HTMLInputElement>;
	dash: HTMLSelectElement;
	effects: HTMLButtonElement;
	patterns: HTMLButtonElement;
}

const NAMES = {
	fill: { legend: 'Fill', none: 'No fill', solid: 'Solid fill' },
	line: { legend: 'Line', none: 'No line', solid: 'Solid line' },
} as const;

/** Visio's dash types 1-23; 0 (no line) is the No line choice. */
export const DASH_TYPES = Array.from({ length: 23 }, (_, index) => ({
	value: String(index + 1),
	label: index === 0 ? 'Solid' : `Dash ${index + 1}`,
}));

function number(doc: Document, field: PaneNumber, text: string, max: number, step: string) {
	const label = doc.createElement('label');
	label.className = 'format-pane-row';
	const caption = doc.createElement('span');
	caption.textContent = text;
	const input = doc.createElement('input');
	input.type = 'number';
	input.min = '0';
	input.max = String(max);
	input.step = step;
	input.dataset.paneField = field;
	label.append(caption, input);
	return { label, input };
}

function section(doc: Document, target: PaneTarget): FormatPaneSection {
	const names = NAMES[target];
	const set = doc.createElement('fieldset');
	set.dataset.formatSection = target;
	const legend = doc.createElement('legend');
	legend.textContent = names.legend;
	const choice = (value: 'none' | 'solid', text: string) => {
		const label = doc.createElement('label');
		label.className = 'format-pane-choice';
		const input = doc.createElement('input');
		input.type = 'radio';
		input.name = `format-${target}`;
		input.value = value;
		label.append(input, text);
		return { label, input };
	};
	const none = choice('none', names.none);
	const solid = choice('solid', names.solid);
	const details = doc.createElement('div');
	details.className = 'format-pane-details';
	const row = doc.createElement('div');
	row.className = 'format-pane-row';
	const caption = doc.createElement('span');
	caption.textContent = 'Color';
	const color = doc.createElement('button');
	color.type = 'button';
	color.className = 'format-pane-color';
	color.dataset.formatColor = target;
	color.setAttribute('aria-label', `${names.legend} color`);
	color.setAttribute('aria-expanded', 'false');
	color.append(doc.createElement('span'));
	row.append(caption, color);
	const grid = colorGrid(doc, target);
	// The pane's own choices say No fill and No line.
	grid.removeAttribute('none-label');
	delete grid.dataset.menuItem;
	grid.hidden = true;
	details.append(row, grid);
	set.append(legend, none.label, solid.label, details);
	return { set, none: none.input, solid: solid.input, color, grid, details };
}

/** The Format Shape task pane view: Fill and Line, as plain controls the pane class drives. */
export function createFormatPaneView(doc: Document): FormatPaneView {
	const root = doc.createElement('section');
	root.className = 'format-pane';
	root.dataset.paneView = 'format';
	const hint = doc.createElement('p');
	hint.className = 'format-pane-hint';
	const fill = section(doc, 'fill');
	const line = section(doc, 'line');
	const fillTransparency = number(doc, 'fillTransparency', 'Transparency (%)', 100, '1');
	const lineTransparency = number(doc, 'lineTransparency', 'Transparency (%)', 100, '1');
	const lineWeight = number(doc, 'lineWeight', 'Width (pt)', 150, '0.25');
	const dashRow = doc.createElement('label');
	dashRow.className = 'format-pane-row';
	const dashCaption = doc.createElement('span');
	dashCaption.textContent = 'Dash type';
	const dash = doc.createElement('select');
	dash.dataset.paneField = 'linePattern';
	dashRow.append(dashCaption, dash);
	fill.details.append(fillTransparency.label);
	line.details.append(lineTransparency.label, lineWeight.label, dashRow);
	const links = doc.createElement('div');
	links.className = 'format-pane-links';
	const link = (text: string, name: string) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.textContent = text;
		button.dataset.formatLink = name;
		return button;
	};
	const effects = link('Effects...', 'effects');
	const patterns = link('Fill and Line Patterns...', 'patterns');
	links.append(effects, patterns);
	root.append(hint, fill.set, line.set, links);
	return {
		section: root,
		hint,
		fill,
		line,
		numbers: {
			fillTransparency: fillTransparency.input,
			lineTransparency: lineTransparency.input,
			lineWeight: lineWeight.input,
		},
		dash,
		effects,
		patterns,
	};
}
