/**
 * The body of `pptx-ui-pdf-export-dialog`: the Range radios (All, Current slide, Slides from-to),
 * "Include hidden slides" and the range error, in the dialog's light DOM so its focus trap sees
 * the controls. Built from the shared `pptx-ui-radio` and `pptx-ui-checkbox`.
 */
import type { PdfExportChoice, PdfExportRange } from '../render/pdf-export-range';
import { parsePdfSlideNumber } from '../render/pdf-export-range';

type Checkable = HTMLElement & { checked: boolean; disabled: boolean; value: string; name: string };

const STYLE = `
.pptx-pdf-export { display: grid; gap: var(--office-space-3, 12px); min-width: 280px; font: var(--office-font-size, 14px)/1.4 var(--office-font, inherit); color: var(--office-foreground, inherit); }
.pptx-pdf-export fieldset { margin: 0; padding: 0; border: 0; display: grid; gap: var(--office-space-2, 8px); }
.pptx-pdf-export legend { padding: 0; margin-bottom: var(--office-space-2, 8px); font-weight: 600; }
.pptx-pdf-export label { display: inline-flex; align-items: center; gap: var(--office-space-2, 8px); cursor: pointer; }
.pptx-pdf-export .pptx-pdf-export-slides { display: flex; flex-wrap: wrap; align-items: center; gap: var(--office-space-2, 8px); }
.pptx-pdf-export input { box-sizing: border-box; width: 5em; min-height: 28px; padding: 2px 6px; border: var(--office-border-width, 1px) solid var(--office-border, #8a8886); border-radius: var(--office-radius, 4px); background: var(--office-background, Field); color: var(--office-foreground, FieldText); font: inherit; }
.pptx-pdf-export input:focus-visible { outline: var(--office-focus-width, 2px) solid var(--office-ring, Highlight); outline-offset: 1px; }
.pptx-pdf-export input:disabled { opacity: .5; }
.pptx-pdf-export .pptx-pdf-export-error { margin: 0; color: var(--office-destructive, #a4262c); }
.pptx-pdf-export .pptx-pdf-export-error:empty { display: none; }
`;

export interface PdfExportDialogText {
	range: string;
	all: string;
	current: string;
	slides: string;
	from: string;
	to: string;
	includeHidden: string;
}

export interface PdfExportDialogBody {
	element: HTMLElement;
	/** Show `choice` in the controls. */
	set(choice: PdfExportChoice, slideCount: number): void;
	/** The choice the controls show. */
	read(): PdfExportChoice;
	label(text: PdfExportDialogText): void;
	error(message: string): void;
}

/** Forward a click on the label text to its control (labels are not form-associated everywhere). */
function forwardLabel(label: HTMLElement, control: HTMLElement): void {
	label.addEventListener('click', (event) => {
		if (event.target === label || (event.target as Node).nodeType === 3) {
			event.preventDefault();
			control.click();
		}
	});
}

export function createPdfExportDialogBody(
	doc: Document,
	onChange: () => void,
	onSubmit: () => void,
): PdfExportDialogBody {
	const element = doc.createElement('div');
	element.className = 'pptx-pdf-export';
	element.dataset.testid = 'pptx-pdf-export-dialog-body';
	const style = doc.createElement('style');
	style.textContent = STYLE;
	const fieldset = doc.createElement('fieldset');
	const legend = doc.createElement('legend');
	fieldset.append(legend);
	const radios = new Map<PdfExportRange, { radio: Checkable; text: Text; row: HTMLElement }>();
	for (const range of ['all', 'current', 'slides'] as const) {
		const row = doc.createElement('div');
		const label = doc.createElement('label');
		const radio = doc.createElement('pptx-ui-radio') as Checkable;
		radio.name = 'pptx-pdf-export-range';
		radio.value = range;
		radio.dataset.testid = `pptx-pdf-export-range-${range}`;
		const text = doc.createTextNode('');
		label.append(radio, text);
		forwardLabel(label, radio);
		row.append(label);
		radio.addEventListener('change', () => sync());
		fieldset.append(row);
		radios.set(range, { radio, text, row });
	}
	const number = (testId: string): HTMLInputElement => {
		const input = doc.createElement('input');
		input.type = 'text';
		input.inputMode = 'numeric';
		input.dataset.testid = testId;
		input.addEventListener('input', () => onChange());
		input.addEventListener('keydown', (event) => {
			if (event.key === 'Enter') {
				event.preventDefault();
				onSubmit();
			}
		});
		return input;
	};
	const from = number('pptx-pdf-export-from');
	const to = number('pptx-pdf-export-to');
	const fromLabel = doc.createElement('label');
	const fromText = doc.createTextNode('');
	fromLabel.append(fromText, from);
	const toLabel = doc.createElement('label');
	const toText = doc.createTextNode('');
	toLabel.append(toText, to);
	const slidesRow = radios.get('slides')!.row;
	slidesRow.className = 'pptx-pdf-export-slides';
	slidesRow.append(fromLabel, toLabel);
	const hiddenLabel = doc.createElement('label');
	const hidden = doc.createElement('pptx-ui-checkbox') as Checkable;
	hidden.dataset.testid = 'pptx-pdf-export-include-hidden';
	const hiddenText = doc.createTextNode('');
	hiddenLabel.append(hidden, hiddenText);
	forwardLabel(hiddenLabel, hidden);
	hidden.addEventListener('change', () => onChange());
	const errorLine = doc.createElement('p');
	errorLine.className = 'pptx-pdf-export-error';
	errorLine.setAttribute('role', 'alert');
	errorLine.dataset.testid = 'pptx-pdf-export-error';
	element.append(style, fieldset, hiddenLabel, errorLine);

	const range = (): PdfExportRange =>
		[...radios].find(([, entry]) => entry.radio.checked)?.[0] ?? 'all';
	function sync(): void {
		const slides = range() === 'slides';
		from.disabled = !slides;
		to.disabled = !slides;
		onChange();
	}

	return {
		element,
		set(choice, slideCount) {
			for (const [id, entry] of radios) entry.radio.checked = id === choice.range;
			from.value = String(choice.from);
			to.value = String(choice.to);
			from.title = to.title = `1-${slideCount}`;
			hidden.checked = choice.includeHiddenSlides;
			sync();
		},
		read() {
			return {
				range: range(),
				from: parsePdfSlideNumber(from.value),
				to: parsePdfSlideNumber(to.value),
				includeHiddenSlides: hidden.checked,
			};
		},
		label(text) {
			legend.textContent = text.range;
			radios.get('all')!.text.textContent = text.all;
			radios.get('current')!.text.textContent = text.current;
			radios.get('slides')!.text.textContent = text.slides;
			for (const [id, entry] of radios)
				entry.radio.setAttribute('aria-label', text[id === 'slides' ? 'slides' : id]);
			fromText.textContent = text.from;
			toText.textContent = text.to;
			from.setAttribute('aria-label', `${text.slides} ${text.from}`);
			to.setAttribute('aria-label', `${text.slides} ${text.to}`);
			hiddenText.textContent = text.includeHidden;
			hidden.setAttribute('aria-label', text.includeHidden);
		},
		error(message) {
			errorLine.textContent = message;
		},
	};
}
