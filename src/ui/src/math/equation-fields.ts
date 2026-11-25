import { latexToOmmlXml, MAX_LATEX_LENGTH } from 'ooxml-core/math';
import { equationPreview, renderEquation } from './equation-render';

export interface EquationDraft {
	omml: string;
	latex: string;
	display: boolean;
}

export interface EquationLabels {
	equation: string;
	source: string;
	display: string;
	preview: string;
	empty: string;
	invalid: string;
}

export function equationMarkup(latex: string, display: boolean): string | undefined {
	try {
		return latexToOmmlXml(latex, display);
	} catch {
		return undefined;
	}
}

/** Shared inputs and live MathML preview, mounted inside the product's existing modal. */
export function createEquationFields(
	doc: Document,
	options: {
		labels: EquationLabels;
		latex?: string;
		display: boolean;
		allowInline?: boolean;
		onChange?(value: EquationDraft | undefined): void;
		accept?(): void;
	},
) {
	const { labels } = options;
	const element = doc.createElement('div');
	const source = doc.createElement('textarea');
	source.rows = 3;
	source.maxLength = MAX_LATEX_LENGTH;
	source.spellcheck = false;
	source.className = 'office-equation-source';
	source.style.boxSizing = 'border-box';
	source.style.width = '100%';
	source.value = options.latex ?? '';
	const sourceLabel = doc.createElement('label');
	sourceLabel.append(labels.source, source);
	element.append(sourceLabel);
	const display = doc.createElement('input');
	display.type = 'checkbox';
	display.checked = options.display;
	if (options.allowInline) {
		const label = doc.createElement('label');
		label.append(display, labels.display);
		element.append(label);
	}
	const preview = doc.createElement('div');
	preview.className = 'office-equation-preview';
	preview.setAttribute('role', 'img');
	preview.setAttribute('aria-label', labels.preview);
	const previewLabel = doc.createElement('span');
	previewLabel.textContent = labels.preview;
	const message = doc.createElement('p');
	message.setAttribute('role', 'status');
	element.append(previewLabel, preview, message);
	let draft: EquationDraft | undefined;
	function refresh() {
		const latex = source.value.trim();
		const omml =
			latex && source.value.length <= MAX_LATEX_LENGTH && equationMarkup(latex, display.checked);
		const drawing = omml ? equationPreview(omml) : undefined;
		const math = drawing && renderEquation(doc, drawing, display.checked, labels.equation);
		preview.replaceChildren(...(math ? [math] : []));
		message.textContent = !latex ? labels.empty : math ? '' : labels.invalid;
		draft = omml && math ? { latex, omml, display: display.checked } : undefined;
		options.onChange?.(draft);
	}
	source.addEventListener('input', refresh);
	display.addEventListener('change', refresh);
	element.addEventListener('keydown', (event) => {
		if (event.key !== 'Enter' || (!event.ctrlKey && !event.metaKey) || !draft) return;
		event.preventDefault();
		options.accept?.();
	});
	refresh();
	return {
		element,
		source,
		preview,
		get draft() {
			return draft;
		},
	};
}
