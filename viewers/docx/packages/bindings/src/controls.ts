import type { WordNavigationProps, WordStatusProps, WordToolbarProps } from 'docx-web-component';

/** One option of a select control. */
export interface ControlOption {
	value: string;
	label: string;
}
interface ControlBase {
	/** Stable across updates, so a binding keeps the control's node (and its focus). */
	id: string;
	label: string;
	disabled: boolean;
}
/** A button that is pressed, not pressed or `mixed` (the selection holds both). */
export interface ToggleControl extends ControlBase {
	kind: 'toggle';
	pressed: boolean | 'mixed';
	run(): void;
}
export interface ActionControl extends ControlBase {
	kind: 'action';
	run(): void;
}
/** A select; `value` is empty when the selection holds more than one value. */
export interface SelectControl extends ControlBase {
	kind: 'select';
	value: string;
	options: readonly ControlOption[];
	change(value: string): void;
}
/** Read-only text such as "Page 2 of 5". */
export interface TextControl {
	kind: 'text';
	id: string;
	label: string;
	text: string;
	disabled: false;
}
export type Control = ToggleControl | ActionControl | SelectControl | TextControl;
/** Controls grouped as the library draws them: one group per cluster of related controls. */
export type ControlGroup = readonly Control[];

const FONT_FAMILIES = [
	'Calibri',
	'Arial',
	'Cambria',
	'Courier New',
	'Georgia',
	'Segoe UI',
	'Times New Roman',
	'Verdana',
];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
const ZOOMS = [50, 75, 100, 125, 150, 200];

/**
 * The listed values plus the current one, which a document may use even when it is unlisted. A
 * mixed selection (`current` null) gets a blank option so the select shows nothing chosen.
 */
function withCurrent(
	values: readonly (string | number)[],
	current: string | number | null,
): ControlOption[] {
	const numeric = typeof values[0] === 'number';
	const all = current === null || values.includes(current) ? [...values] : [...values, current];
	const sorted = all
		.map(String)
		.sort((left, right) => (numeric ? Number(left) - Number(right) : left.localeCompare(right)))
		.map((value) => ({ value, label: value }));
	return current === null ? [{ value: '', label: '' }, ...sorted] : sorted;
}

const label = (labels: Record<string, string>, key: string) => labels[key] ?? key;

/** Controls for the library Toolbar. Each reads its state from `props` and runs its callback. */
export function toolbarControls(props: WordToolbarProps): ControlGroup[] {
	const { labels } = props;
	const edit = !props.canEdit;
	const toggle = (id: string, key: string, pressed: boolean | null, run: () => unknown) =>
		({
			kind: 'toggle',
			id,
			label: label(labels, key),
			disabled: edit,
			pressed: pressed === null ? 'mixed' : pressed,
			run: () => void run(),
		}) satisfies ToggleControl;
	const action = (id: string, key: string, disabled: boolean, run: () => unknown) =>
		({
			kind: 'action',
			id,
			label: label(labels, key),
			disabled,
			run: () => void run(),
		}) satisfies ActionControl;
	const alignment = (id: string, key: string, value: typeof props.alignment & string) =>
		({
			kind: 'toggle',
			id,
			label: label(labels, key),
			disabled: edit,
			pressed: props.alignment === null ? 'mixed' : props.alignment === value,
			run: () => void props.onAlignmentChange(value),
		}) satisfies ToggleControl;
	const list = (id: string, key: string, kind: 'bullet' | 'number') =>
		({
			kind: 'toggle',
			id,
			label: label(labels, key),
			disabled: edit,
			pressed: props.list === kind,
			run: () => void props.onListChange(props.list === kind ? 'remove' : kind),
		}) satisfies ToggleControl;
	const groups: Control[][] = [
		[
			action('undo', 'Undo', !props.canUndo, props.onUndo),
			action('redo', 'Redo', !props.canRedo, props.onRedo),
		],
		[
			{
				kind: 'select',
				id: 'fontFamily',
				label: label(labels, 'Font'),
				disabled: edit,
				value: props.fontFamily ?? '',
				options: withCurrent(FONT_FAMILIES, props.fontFamily),
				change: (value) => void props.onFontFamilyChange(value),
			},
			{
				kind: 'select',
				id: 'fontSize',
				label: label(labels, 'Font size'),
				disabled: edit,
				value: props.fontSize === null ? '' : String(props.fontSize),
				options: withCurrent(FONT_SIZES, props.fontSize),
				change: (value) => void props.onFontSizeChange(Number(value)),
			},
		],
		[
			toggle('bold', 'Bold', props.bold, props.onBold),
			toggle('italic', 'Italic', props.italic, props.onItalic),
			toggle('underline', 'Underline', props.underline, props.onUnderline),
		],
		[
			alignment('alignLeft', 'Align left', 'left'),
			alignment('alignCenter', 'Align center', 'center'),
			alignment('alignRight', 'Align right', 'right'),
			alignment('alignJustify', 'Justify', 'justify'),
		],
		[list('bulletList', 'Bulleted list', 'bullet'), list('numberList', 'Numbered list', 'number')],
	];
	// Host workflows appear only when the host supplied them; they are disabled while unusable.
	const host: Control[] = [];
	const { hyperlink, findReplace, picture } = props.hostWorkflows;
	if (hyperlink)
		host.push(action('link', 'Insert link', !props.canInsertLink, props.openHyperlinkDialog));
	if (findReplace) host.push(action('find', 'Find', !props.canFindReplace, props.openFindReplace));
	if (picture)
		host.push(action('picture', 'Insert picture', !props.canInsertPicture, props.openImagePicker));
	if (host.length) groups.push(host);
	return groups;
}

/** Controls for the library PageNavigation: previous, page text, next and the layout switch. */
export function navigationControls(props: WordNavigationProps): ControlGroup[] {
	const { labels } = props;
	return [
		[
			{
				kind: 'action',
				id: 'previousPage',
				label: label(labels, 'Previous page'),
				disabled: !props.canGoPrevious,
				run: () => void props.onPreviousPage(),
			},
			{ kind: 'text', id: 'pageText', label: '', text: props.pageText, disabled: false },
			{
				kind: 'action',
				id: 'nextPage',
				label: label(labels, 'Next page'),
				disabled: !props.canGoNext,
				run: () => void props.onNextPage(),
			},
		],
		[
			{
				kind: 'select',
				id: 'viewMode',
				label: label(labels, 'Layout view'),
				disabled: false,
				value: props.viewMode,
				options: [
					{ value: 'draft', label: label(labels, 'Draft') },
					{ value: 'print', label: label(labels, 'Print Layout') },
				],
				change: (value) => void props.onViewModeChange(value === 'print' ? 'print' : 'draft'),
			},
		],
	];
}

/** Controls for the library StatusBar: word count and zoom. */
export function statusControls(props: WordStatusProps): ControlGroup[] {
	const zooms = withCurrent(ZOOMS, props.zoom).map((option) => ({
		value: option.value,
		label: `${option.label}%`,
	}));
	return [
		[{ kind: 'text', id: 'wordText', label: '', text: props.wordText, disabled: false }],
		[
			{
				kind: 'select',
				id: 'zoom',
				label: label(props.labels, 'Zoom'),
				disabled: false,
				value: String(props.zoom),
				options: zooms,
				change: (value) => void props.onZoomChange(Number(value)),
			},
		],
	];
}
