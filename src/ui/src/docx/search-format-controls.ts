// Format > Font, Paragraph, and Style picker UI for Word's Find & Replace search panel.
import type {
	FindFormatCriteria,
	ParagraphAlign,
	ReplaceFormatPatch,
	StHighlightColor,
} from 'ooxml-core/docx';
import { hasFormatCriteria, hasReplaceFormat } from 'ooxml-core/docx';
import { formatSummary } from './search-format-match';
import { translate, type EditorLocale } from './localization';

const DEFAULT_FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Courier New'];
const DEFAULT_STYLES = [
	'Normal',
	'Heading1',
	'Heading2',
	'Heading3',
	'Heading4',
	'Title',
	'Subtitle',
	'Quote',
];
const HIGHLIGHTS: readonly StHighlightColor[] = [
	'yellow',
	'green',
	'cyan',
	'magenta',
	'blue',
	'red',
	'none',
];
const ALIGNMENTS: readonly ParagraphAlign[] = ['left', 'center', 'right', 'justify'];

export interface FormatPickerHandle<T> {
	readonly element: HTMLElement;
	readonly summary: HTMLElement;
	value(): T | undefined;
	clear(): void;
	refreshSummary(locale: EditorLocale): void;
}

function makeSelect(
	label: string,
	options: ReadonlyArray<readonly [string, string]>,
	ariaLabel: string,
): HTMLSelectElement {
	const el = document.createElement('select');
	el.setAttribute('aria-label', ariaLabel);
	el.dataset.formatField = label;
	for (const [value, text] of options) {
		const opt = document.createElement('option');
		opt.value = value;
		opt.textContent = text;
		el.append(opt);
	}
	return el;
}

function triState(name: string, label: string): HTMLSelectElement {
	return makeSelect(
		name,
		[
			['', `${label}: Any`],
			['on', label],
			['off', `Not ${label}`],
		],
		label,
	);
}

/** Builds one Format > Font / Paragraph / Style fieldset for either Find or Replace. */
export function createFormatPicker<T extends FindFormatCriteria | ReplaceFormatPatch>(
	kind: 'find' | 'replace',
	onChange: () => void,
	getStyles?: () => readonly string[],
): FormatPickerHandle<T> {
	const prefix = kind === 'find' ? 'Find format' : 'Replace format';
	const fieldset = document.createElement('fieldset');
	fieldset.className = 'dve-search-format';
	fieldset.dataset.formatKind = kind;

	const legend = document.createElement('legend');
	legend.textContent = prefix;

	const fontSelect = makeSelect(
		'fontFamily',
		[['', 'Font: Any'], ...DEFAULT_FONTS.map((f) => [f, f] as const)],
		`${prefix} font`,
	);
	const sizeInput = document.createElement('input');
	sizeInput.type = 'number';
	sizeInput.min = '4';
	sizeInput.max = '1638';
	sizeInput.step = '0.5';
	sizeInput.placeholder = 'pt';
	sizeInput.className = 'dve-search-size';
	sizeInput.setAttribute('aria-label', `${prefix} size`);
	sizeInput.dataset.formatField = 'fontSizePt';

	const boldSelect = triState('bold', 'Bold');
	const italicSelect = triState('italic', 'Italic');
	const underlineSelect = triState('underline', 'Underline');

	const colorInput = document.createElement('input');
	colorInput.type = 'text';
	colorInput.placeholder = '#RRGGBB';
	colorInput.className = 'dve-search-color';
	colorInput.setAttribute('aria-label', `${prefix} color`);
	colorInput.dataset.formatField = 'color';

	const highlightSelect = makeSelect(
		'highlight',
		[['', 'Highlight: Any'], ...HIGHLIGHTS.map((h) => [h, h] as const)],
		`${prefix} highlight`,
	);

	const alignSelect = makeSelect(
		'alignment',
		[['', 'Alignment: Any'], ...ALIGNMENTS.map((a) => [a, a] as const)],
		`${prefix} alignment`,
	);

	const styleIds = [...new Set([...DEFAULT_STYLES, ...(getStyles?.() ?? [])])];
	const styleSelect = makeSelect(
		'styleId',
		[['', 'Style: Any'], ...styleIds.map((s) => [s, s] as const)],
		`${prefix} style`,
	);

	const clearBtn = document.createElement('button');
	clearBtn.type = 'button';
	clearBtn.textContent = 'No Formatting';
	clearBtn.setAttribute('aria-label', `${prefix}: No Formatting`);

	const summary = document.createElement('span');
	summary.className = 'dve-search-format-summary';
	summary.dataset.formatSummary = kind;

	const read = (): T | undefined => {
		const out: Record<string, unknown> = {};
		if (fontSelect.value) out.fontFamily = fontSelect.value;
		const size = Number.parseFloat(sizeInput.value);
		if (Number.isFinite(size) && size > 0) out.fontSizePt = size;
		if (boldSelect.value === 'on') out.bold = true;
		else if (boldSelect.value === 'off') out.bold = false;
		if (italicSelect.value === 'on') out.italic = true;
		else if (italicSelect.value === 'off') out.italic = false;
		if (underlineSelect.value === 'on') out.underline = true;
		else if (underlineSelect.value === 'off') out.underline = false;
		const color = colorInput.value.trim();
		if (/^#?[0-9a-f]{6}$/i.test(color)) out.color = color;
		if (highlightSelect.value) out.highlight = highlightSelect.value as StHighlightColor;
		if (alignSelect.value) out.alignment = alignSelect.value as ParagraphAlign;
		if (styleSelect.value) out.styleId = styleSelect.value;
		const hasAny =
			kind === 'find'
				? hasFormatCriteria(out as FindFormatCriteria)
				: hasReplaceFormat(out as ReplaceFormatPatch);
		return hasAny ? (out as T) : undefined;
	};

	const clear = () => {
		fontSelect.value = '';
		sizeInput.value = '';
		boldSelect.value = '';
		italicSelect.value = '';
		underlineSelect.value = '';
		colorInput.value = '';
		highlightSelect.value = '';
		alignSelect.value = '';
		styleSelect.value = '';
		onChange();
	};

	for (const control of [
		fontSelect,
		boldSelect,
		italicSelect,
		underlineSelect,
		highlightSelect,
		alignSelect,
		styleSelect,
	]) {
		control.addEventListener('change', onChange);
	}
	for (const input of [sizeInput, colorInput]) {
		input.addEventListener('input', onChange);
	}
	clearBtn.addEventListener('click', clear);

	fieldset.append(
		legend,
		fontSelect,
		sizeInput,
		boldSelect,
		italicSelect,
		underlineSelect,
		colorInput,
		highlightSelect,
		alignSelect,
		styleSelect,
		clearBtn,
		summary,
	);

	return {
		element: fieldset,
		summary,
		value: read,
		clear,
		refreshSummary(locale: EditorLocale) {
			const text = formatSummary(locale, read());
			summary.textContent = text ? `${translate(locale, 'Format')}: ${text}` : '';
		},
	};
}
