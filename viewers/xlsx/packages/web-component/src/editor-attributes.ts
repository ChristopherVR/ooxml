import { normalizeEditorLocale } from './localization';
import { normalizeThemeMode, type EditorThemeMode } from './theme';

/**
 * Attribute <-> property mapping for `<xlsx-editor>`. Properties are the source of truth: an
 * attribute change goes through the property setter only when it would change the value, and a
 * property change is written back only when the attribute would differ, so the two directions
 * cannot ping-pong (the docx editor's scheme).
 */
export const XLSX_EDITOR_ATTRIBUTES = [
	'locale',
	'read-only',
	'file-name',
	'author-name',
	'theme',
	'show-toolbar',
	'show-formula-bar',
] as const;
export type XlsxEditorAttribute = (typeof XLSX_EDITOR_ATTRIBUTES)[number];

export const DEFAULT_FILE_NAME = 'Book1.xlsx';
export const DEFAULT_AUTHOR_NAME = 'Author';

/** The element properties the attributes drive. */
export interface AttributeProperties {
	locale: string;
	readOnly: boolean;
	fileName: string;
	authorName: string;
	theme: EditorThemeMode;
	showToolbar: boolean;
	showFormulaBar: boolean;
}

export function isXlsxEditorAttribute(name: string): name is XlsxEditorAttribute {
	return (XLSX_EDITOR_ATTRIBUTES as readonly string[]).includes(name);
}

/** attributeChangedCallback body: attribute -> property. */
export function applyAttribute(
	target: AttributeProperties,
	name: XlsxEditorAttribute,
	value: string | null,
): void {
	switch (name) {
		case 'locale':
			if (normalizeEditorLocale(value ?? 'en') !== target.locale) target.locale = value ?? 'en';
			return;
		case 'read-only':
			if ((value !== null && value !== 'false') !== target.readOnly)
				target.readOnly = value !== null && value !== 'false';
			return;
		case 'file-name': {
			const next = value || DEFAULT_FILE_NAME;
			if (next !== target.fileName) target.fileName = next;
			return;
		}
		case 'author-name': {
			const next = value || DEFAULT_AUTHOR_NAME;
			if (next !== target.authorName) target.authorName = next;
			return;
		}
		case 'theme': {
			const next = normalizeThemeMode(value);
			if (next !== target.theme) target.theme = next;
			return;
		}
		case 'show-toolbar':
			if ((value !== 'false') !== target.showToolbar) target.showToolbar = value !== 'false';
			return;
		case 'show-formula-bar':
			if ((value !== 'false') !== target.showFormulaBar) target.showFormulaBar = value !== 'false';
	}
}

/** Property setter tail: property -> attribute. A no-op where the element has no DOM (SSR). */
export function reflectAttribute(
	element: Element,
	name: XlsxEditorAttribute,
	value: string | boolean,
): void {
	if (typeof element.setAttribute !== 'function') return;
	if ((name === 'show-toolbar' || name === 'show-formula-bar') && typeof value === 'boolean') {
		// Default true: the attribute only appears to opt out.
		if (element.getAttribute(name) !== (value ? null : 'false')) {
			if (value) element.removeAttribute(name);
			else element.setAttribute(name, 'false');
		}
		return;
	}
	if (typeof value === 'boolean') {
		if (element.hasAttribute(name) !== value) element.toggleAttribute(name, value);
		return;
	}
	const current = element.getAttribute(name);
	if (current === value) return;
	// A differently spelled tag such as "FR" already means this locale; leave the author's markup.
	if (name === 'locale' && current !== null && normalizeEditorLocale(current) === value) return;
	element.setAttribute(name, value);
}
