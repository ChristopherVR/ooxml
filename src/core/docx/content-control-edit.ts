// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// DOM-free editing of content controls: what a click on a checkbox, a pick from a list or a date
// chosen in the calendar does to the control and to the text it shows.
import { formatWordDate } from './field-date';
import { isContentControlLocked, type ContentControl } from './content-control-model';

/** A control's new state and the text (with an optional font) its content should now show. */
export interface ControlChange {
	control: ContentControl;
	text: string;
	font?: string;
	/** The control is temporary (`w:temporary`): the edit keeps `text` and removes the control. */
	removed?: boolean;
}

const DEFAULT_GLYPHS = {
	checked: { code: '2612', font: 'MS Gothic' },
	unchecked: { code: '2610', font: 'MS Gothic' },
};
const glyphText = (code: string) => String.fromCodePoint(Number.parseInt(code, 16) || 0x2610);

/** Whether the user may type in, toggle or pick for this control. */
export const canEditControlContent = (control: ContentControl | undefined): boolean =>
	!isContentControlLocked(control, 'content');
/** Whether the control itself may be deleted (with its content). */
export const canDeleteControl = (control: ContentControl | undefined): boolean =>
	!isContentControlLocked(control, 'control');

/**
 * Whether editing the control's content removes it, its content staying as ordinary text: Word
 * removes a temporary control (ECMA-376 17.5.2.43) unless its lock forbids deleting it.
 */
export const removedByEdit = (control: ContentControl): boolean =>
	Boolean(control.temporary) && canDeleteControl(control);

/** The change as Word applies it (`removedByEdit`). */
export const afterEdit = (change: ControlChange): ControlChange =>
	removedByEdit(change.control) ? { ...change, removed: true } : change;

export function checkboxState(control: ContentControl, checked: boolean): ControlChange {
	const glyph =
		(checked ? control.checkedGlyph : control.uncheckedGlyph) ??
		(checked ? DEFAULT_GLYPHS.checked : DEFAULT_GLYPHS.unchecked);
	return {
		control: { ...control, checked },
		text: glyphText(glyph.code),
		...(glyph.font ? { font: glyph.font } : {}),
	};
}

export const toggleCheckbox = (control: ContentControl): ControlChange =>
	afterEdit(checkboxState(control, !control.checked));

/** Choosing entry `index` of a drop-down list or combo box. */
export function selectListItem(control: ContentControl, index: number): ControlChange | undefined {
	const item = control.items?.[index];
	if (!item) return undefined;
	const { showingPlaceholder: _placeholder, ...rest } = control;
	return afterEdit({ control: rest, text: item.display });
}

/** The date the control holds, from `w:fullDate`. */
export function controlDate(control: ContentControl): Date | undefined {
	if (!control.fullDate) return undefined;
	const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(control.fullDate);
	if (!match) return undefined;
	return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

/** `w:fullDate` for a calendar day: local midnight written as UTC-less ISO, as Word does. */
export const fullDateOf = (date: Date): string =>
	`${String(date.getFullYear()).padStart(4, '0')}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T00:00:00Z`;

/** Choosing `date` in the calendar: shown with the control's date format. */
export function setControlDate(control: ContentControl, date: Date): ControlChange {
	const { showingPlaceholder: _placeholder, ...rest } = control;
	return afterEdit({
		control: { ...rest, fullDate: fullDateOf(date) },
		text: formatWordDate(date, control.dateFormat ?? 'M/d/yyyy', control.dateLocale ?? 'en-US'),
	});
}

/** The placeholder text Word shows for an empty control of this kind. */
export function defaultPlaceholder(control: ContentControl): string {
	switch (control.kind) {
		case 'date':
			return 'Click or tap to enter a date.';
		case 'dropDownList':
		case 'comboBox':
			return 'Choose an item.';
		case 'picture':
			return '';
		default:
			return 'Click or tap here to enter text.';
	}
}

/**
 * The control's own placeholder text (from the glossary), else Word's default for its kind. An
 * entry without text gives the default too: an inline control left without text is lost with it.
 */
export const placeholderTextOf = (control: ContentControl): string =>
	control.placeholderText || defaultPlaceholder(control);

/** Typed text replaces a placeholder: the control no longer shows one. */
export function withValue(control: ContentControl): ContentControl {
	const { showingPlaceholder: _placeholder, ...rest } = control;
	return rest;
}
