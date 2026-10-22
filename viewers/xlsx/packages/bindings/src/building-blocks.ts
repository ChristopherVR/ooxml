import { deferredHandle, mountEditorSurface, type SurfaceOptions } from './index';
import type { SpreadsheetToolbarProps, XlsxCollaborationState } from 'xlsx-web-component';

export type BuildingBlockOptions = SurfaceOptions;
export interface BuildingBlockState {
	toolbarProps: SpreadsheetToolbarProps | null;
	loading: boolean;
	error: Error | null;
	mode: 'view' | 'edit';
	/** Unsaved changes since the last load, save or `markClean()`. */
	dirty: boolean;
	/** The selection as A1 text. */
	selection: string;
	/** Sharing state of the open workbook. */
	collaboration: XlsxCollaborationState;
	/** The validation message of the cell edit in progress, or null. */
	editError: string | null;
}
export const OFF_COLLABORATION: XlsxCollaborationState = Object.freeze({
	active: false,
	roomId: undefined,
	status: 'off',
	synced: false,
	people: [],
});
/** The reactive values every framework returns next to `toolbarProps`, derived from them. */
export function derivedBlockState(
	props: SpreadsheetToolbarProps | null,
): Pick<BuildingBlockState, 'dirty' | 'selection' | 'collaboration' | 'editError'> {
	return {
		dirty: props?.dirty ?? false,
		selection: props?.selection ?? '',
		collaboration: props?.collaboration ?? OFF_COLLABORATION,
		editError: props?.editError ?? null,
	};
}
/** Framework-neutral engine ownership. Framework adapters own their reactive state and lifecycle. */
export function createBuildingBlockSession(changed: (state: BuildingBlockState) => void) {
	let currentHost: HTMLElement | null = null;
	let binding: ReturnType<typeof mountEditorSurface> | undefined;
	let stop: (() => void) | undefined;
	let options: BuildingBlockOptions = {};
	let error: Error | null = null;
	let wasLoading = false;
	const refresh = () => {
		const loading = binding?.element.loading ?? false;
		if (loading && !wasLoading) error = null;
		wasLoading = loading;
		const toolbarProps = binding?.getToolbarProps() ?? null;
		changed({
			toolbarProps,
			loading: binding?.element.loading ?? false,
			error,
			mode: options.canEdit !== true || options.readOnly ? 'view' : 'edit',
			...derivedBlockState(toolbarProps),
		});
	};
	const effective = (): BuildingBlockOptions => ({
		...options,
		onWorkbookError: (cause) => {
			error = cause;
			wasLoading = binding?.element.loading ?? false;
			refresh();
			options.onWorkbookError?.(cause);
		},
	});
	const unmount = () => {
		currentHost = null;
		stop?.();
		stop = undefined;
		binding?.destroy();
		binding = undefined;
		refresh();
	};
	return {
		handle: deferredHandle(() => binding),
		mount(host: HTMLElement | null) {
			if (host && host === currentHost) return;
			unmount();
			if (!host) return;
			error = null;
			binding = mountEditorSurface(host, effective());
			currentHost = host;
			stop = binding.subscribe(refresh);
			refresh();
		},
		update(next: BuildingBlockOptions) {
			if (
				next.workbook !== options.workbook ||
				next.bytes !== options.bytes ||
				next.src !== options.src
			)
				error = null;
			options = next;
			binding?.update(effective());
			refresh();
		},
		destroy: unmount,
	};
}
export const emptyBuildingBlockState = (): BuildingBlockState => ({
	toolbarProps: null,
	loading: false,
	error: null,
	mode: 'view',
	...derivedBlockState(null),
});
export interface ToolbarButtonModel {
	kind: 'button';
	id: string;
	label: string;
	/** true or false, or null for a mixed selection. */
	pressed: boolean | null;
	disabled: boolean;
	run: () => Promise<boolean>;
}
export interface ToolbarSelectModel {
	kind: 'select';
	id: string;
	label: string;
	/** The current value; '' when the selection is mixed. */
	value: string;
	disabled: boolean;
	options: { value: string; label: string }[];
	change: (value: string) => Promise<boolean>;
}
export type ToolbarItemModel = ToolbarButtonModel | ToolbarSelectModel;
const withCurrent = (options: { value: string; label: string }[], current: string | null) => {
	if (current === null) return [{ value: '', label: '' }, ...options];
	return options.some((o) => o.value === current)
		? options
		: [{ value: current, label: current }, ...options];
};
/** Every control the props support, in display order, for a framework to map onto its template. */
export function toolbarModel(p: SpreadsheetToolbarProps): ToolbarItemModel[] {
	const label = (key: string) => p.labels[key] ?? key;
	const can = (flag: boolean) => !p.canEdit || !flag;
	const button = (
		id: string,
		text: string,
		pressed: boolean | null,
		disabled: boolean,
		run: () => Promise<boolean>,
	): ToolbarButtonModel => ({ kind: 'button', id, label: text, pressed, disabled, run });
	return [
		button('undo', label('Undo'), false, !p.enabled.undo, p.onUndo),
		button('redo', label('Redo'), false, !p.enabled.redo, p.onRedo),
		{
			kind: 'select',
			id: 'fontFamily',
			label: label('Font'),
			value: p.fontFamily ?? '',
			disabled: can(p.enabled.fontFamily),
			options: withCurrent(
				p.options.fontFamilies.map((f) => ({ value: f, label: f })),
				p.fontFamily,
			),
			change: (value) => p.onFontFamilyChange(value),
		},
		{
			kind: 'select',
			id: 'fontSize',
			label: label('Font Size'),
			value: p.fontSize === null ? '' : String(p.fontSize),
			disabled: can(p.enabled.fontSize),
			options: withCurrent(
				p.options.fontSizes.map((n) => ({ value: String(n), label: String(n) })),
				p.fontSize === null ? null : String(p.fontSize),
			),
			change: (value) => p.onFontSizeChange(Number(value)),
		},
		button('bold', label('Bold'), p.bold, can(p.enabled.bold), p.onBold),
		button('italic', label('Italic'), p.italic, can(p.enabled.italic), p.onItalic),
		button('underline', label('Underline'), p.underline, can(p.enabled.underline), p.onUnderline),
		...p.options.alignments.map((a) =>
			button(
				`align-${a.value}`,
				a.label,
				p.alignment === null ? null : p.alignment === a.value,
				can(p.enabled.alignment),
				() => p.onAlignmentChange(a.value),
			),
		),
		{
			kind: 'select',
			id: 'numberFormat',
			label: label('Number Format'),
			value: p.numberFormat ?? '',
			disabled: can(p.enabled.numberFormat),
			options: withCurrent(p.options.numberFormats, p.numberFormat),
			change: (value) => p.onNumberFormatChange(value),
		},
	];
}
export function selectReference(p: SpreadsheetToolbarProps, value: string): boolean {
	try {
		return p.onSelect(value);
	} catch {
		return false;
	}
}
export type { SpreadsheetToolbarProps } from 'xlsx-web-component';
