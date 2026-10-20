import {
	createElement,
	useCallback,
	useEffect,
	useImperativeHandle,
	useRef,
	useState,
	type Ref,
} from 'react';
import {
	mountEditor,
	deferredHandle,
	type EditorOptions,
	type EditorBinding,
	type EditorHandle,
} from './index';
import type { SpreadsheetToolbarProps } from 'xlsx-web-component';
export interface BuildingBlockOptions extends EditorOptions {
	canEdit?: boolean | undefined;
	handle?: Ref<EditorHandle> | undefined;
}
export interface CanvasProps {
	mount: (host: HTMLDivElement | null) => void;
	className?: string | undefined;
}
/** A document surface with no application controls. Keep the returned props together. */
export function SpreadsheetCanvas(props: CanvasProps) {
	return createElement('div', { ref: props.mount, className: props.className });
}
export function useViewerBuildingBlocks(options: BuildingBlockOptions = {}) {
	const latest = useRef(options);
	latest.current = options;
	const binding = useRef<EditorBinding | null>(null);
	const stop = useRef<(() => void) | undefined>(undefined);
	const [toolbarProps, setToolbarProps] = useState<SpreadsheetToolbarProps | null>(null);
	const [loading, setLoading] = useState(false);
	const [mountVersion, setMountVersion] = useState(0);
	const [error, setError] = useState<Error | null>(null);
	const lastBytes = useRef<Uint8Array | ArrayBuffer | undefined>(undefined);
	const effective = () => ({
		...latest.current,
		bytes: undefined,
		readOnly: latest.current.canEdit === false || latest.current.readOnly === true,
		onWorkbookError: (error: Error) => {
			setError(error);
			latest.current.onWorkbookError?.(error);
		},
	});
	const refresh = () => {
		setLoading(binding.current?.element.loading ?? false);
		if (binding.current?.element.loading) setError(null);
		const next = binding.current?.element.getToolbarProps() ?? null;
		setToolbarProps((previous) =>
			JSON.stringify(previous) === JSON.stringify(next) ? previous : next,
		);
	};
	const mount = useCallback((host: HTMLDivElement | null) => {
		stop.current?.();
		stop.current = undefined;
		binding.current?.destroy();
		binding.current = null;
		if (!host) {
			setToolbarProps(null);
			setLoading(false);
			return;
		}
		setMountVersion((version) => version + 1);
		lastBytes.current = undefined;
		binding.current = mountEditor(host, effective(), true);
		stop.current = binding.current.element.subscribeEditorState(refresh);
		refresh();
	}, []);
	useEffect(() => {
		setError(null);
	}, [options.workbook]);
	useEffect(() => {
		binding.current?.update(effective());
	}, [options]);

	useEffect(() => {
		const bytes = options.bytes;
		if (!bytes || !binding.current || lastBytes.current === bytes) return;
		lastBytes.current = bytes;
		const mounted = binding.current;
		const active = () => binding.current === mounted && lastBytes.current === bytes;
		setLoading(true);
		setError(null);
		void binding.current
			.load(bytes)
			.catch((cause: unknown) => {
				if (active()) setError(cause instanceof Error ? cause : new Error(String(cause)));
			})
			.finally(() => {
				if (active()) {
					setLoading(false);
					refresh();
				}
			});
	}, [options.bytes, mountVersion]);
	useImperativeHandle(options.handle, () => deferredHandle(() => binding.current), []);
	return {
		canvasProps: { mount } satisfies CanvasProps,
		toolbarProps,
		loading,
		error,
		mode: options.canEdit === false || options.readOnly ? ('view' as const) : ('edit' as const),
	};
}
export type { SpreadsheetToolbarProps } from 'xlsx-web-component';

/** Optional controls using exactly the callbacks exposed to custom host controls. */
export function Toolbar(props: SpreadsheetToolbarProps) {
	const buttons = [
		['Bold', props.bold, props.onBold, !props.canEdit || !props.enabled.bold],
		['Italic', props.italic, props.onItalic, !props.canEdit || !props.enabled.italic],
		['Underline', props.underline, props.onUnderline, !props.canEdit || !props.enabled.underline],
		['Undo', false, props.onUndo, !props.canUndo],
		['Redo', false, props.onRedo, !props.canRedo],
	] as const;
	return createElement(
		'div',
		{ role: 'toolbar' },
		...buttons.map(([label, pressed, run, disabled]) =>
			createElement(
				'button',
				{
					key: label,
					type: 'button',
					disabled,
					'aria-pressed': pressed === null ? 'mixed' : pressed,
					onMouseDown: (event: import('react').MouseEvent) => event.preventDefault(),
					onClick: () => {
						void run();
					},
				},
				props.labels[label] ?? label,
			),
		),
	);
}

export function FormulaBar(props: SpreadsheetToolbarProps) {
	const [invalid, setInvalid] = useState(false);
	const [invalidName, setInvalidName] = useState(false);
	const input = useRef<HTMLInputElement>(null);
	useEffect(() => {
		if (input.current && input.current.value !== props.formula) input.current.value = props.formula;
	}, [props.formula, props.activeCell]);
	return createElement(
		'div',
		null,
		createElement('input', {
			key: props.selection,
			'aria-label': props.labels['Name Box'],
			defaultValue: props.selection,
			'aria-invalid': invalidName,
			onChange: () => setInvalidName(false),
			onKeyDown: (event: import('react').KeyboardEvent<HTMLInputElement>) => {
				if (event.key === 'Enter') {
					event.preventDefault();
					try {
						setInvalidName(!props.onSelect(event.currentTarget.value));
					} catch {
						setInvalidName(true);
					}
				}
			},
		}),
		createElement('input', {
			key: props.activeCell,
			ref: input,
			'aria-label': props.labels['Formula Bar'],
			defaultValue: props.formula,
			readOnly: !props.canEdit,
			'aria-invalid': invalid,
			onChange: (event: import('react').ChangeEvent<HTMLInputElement>) => {
				setInvalid(false);
				props.onFormulaInput(
					event.currentTarget.value,
					event.currentTarget.selectionEnd ?? undefined,
				);
			},
			onKeyDown: (event: import('react').KeyboardEvent<HTMLInputElement>) => {
				if (event.key === 'Enter') {
					event.preventDefault();
					setInvalid(!props.onFormulaChange(event.currentTarget.value));
				} else if (event.key === 'Escape') {
					event.preventDefault();
					props.onCancelEdit();
					setInvalid(false);
				}
			},
		}),
	);
}
export function SheetTabs(props: SpreadsheetToolbarProps) {
	return createElement(
		'div',
		{ role: 'tablist', 'aria-label': props.labels['Sheet'] },
		...props.sheets
			.filter((sheet) => sheet.state === 'visible')
			.map((sheet) =>
				createElement(
					'button',
					{
						key: sheet.index,
						type: 'button',
						role: 'tab',
						'aria-selected': props.activeSheet === sheet.index,
						onClick: () => props.onSheetChange(sheet.index),
					},
					sheet.name,
				),
			),
	);
}
