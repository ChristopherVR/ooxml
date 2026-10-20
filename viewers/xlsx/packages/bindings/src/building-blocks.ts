import { deferredHandle, mountEditorSurface, type SurfaceOptions } from './index';
import type { SpreadsheetToolbarProps } from 'xlsx-web-component';

export type BuildingBlockOptions = SurfaceOptions;
export interface BuildingBlockState {
	toolbarProps: SpreadsheetToolbarProps | null;
	loading: boolean;
	error: Error | null;
	mode: 'view' | 'edit';
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
		changed({
			toolbarProps: binding?.getToolbarProps() ?? null,
			loading: binding?.element.loading ?? false,
			error,
			mode: options.canEdit === false || options.readOnly ? 'view' : 'edit',
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
	mode: 'edit',
});
export function toolbarButtons(p: SpreadsheetToolbarProps) {
	return [
		{
			label: p.labels['Bold'] ?? 'Bold',
			pressed: p.bold,
			disabled: !p.canEdit || !p.enabled.bold,
			run: p.onBold,
		},
		{
			label: p.labels['Italic'] ?? 'Italic',
			pressed: p.italic,
			disabled: !p.canEdit || !p.enabled.italic,
			run: p.onItalic,
		},
		{
			label: p.labels['Underline'] ?? 'Underline',
			pressed: p.underline,
			disabled: !p.canEdit || !p.enabled.underline,
			run: p.onUnderline,
		},
		{ label: p.labels['Undo'] ?? 'Undo', pressed: false, disabled: !p.canUndo, run: p.onUndo },
		{ label: p.labels['Redo'] ?? 'Redo', pressed: false, disabled: !p.canRedo, run: p.onRedo },
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
