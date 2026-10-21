import { prepareBlock } from '../../bindings/src/block-theme';
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
	mountEditorSurface,
	deferredHandle,
	type SurfaceOptions,
	type EditorHandle,
} from '../../bindings/src/index';
import { derivedBlockState, toolbarModel } from '../../bindings/src/building-blocks';
import type { SpreadsheetToolbarProps } from 'xlsx-web-component';
type Surface = ReturnType<typeof mountEditorSurface>;
export interface BuildingBlockOptions extends SurfaceOptions {
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
	const binding = useRef<Surface | null>(null);
	const stop = useRef<(() => void) | undefined>(undefined);
	const [toolbarProps, setToolbarProps] = useState<SpreadsheetToolbarProps | null>(null);
	const [loading, setLoading] = useState(false);
	const [mountVersion, setMountVersion] = useState(0);
	const snapshotElement = useRef<Surface['element'] | undefined>(undefined);
	const [error, setError] = useState<Error | null>(null);
	const lastBytes = useRef<Uint8Array | ArrayBuffer | undefined>(undefined);
	const effective = () => ({
		...latest.current,
		bytes: undefined,
		readOnly: latest.current.canEdit !== true || latest.current.readOnly === true,
		onWorkbookError: (cause: Error) => {
			setError(cause);
			latest.current.onWorkbookError?.(cause);
		},
	});
	const refresh = () => {
		setLoading(binding.current?.element.loading ?? false);
		if (binding.current?.element.loading) setError(null);
		const element = binding.current?.element;
		const sameElement = snapshotElement.current === element;
		snapshotElement.current = element;
		const next = element?.getToolbarProps() ?? null;
		setToolbarProps((previous) =>
			sameElement && JSON.stringify(previous) === JSON.stringify(next) ? previous : next,
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
		binding.current = mountEditorSurface(host, effective());
		stop.current = binding.current.subscribe(refresh);
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
		...derivedBlockState(toolbarProps),
		loading,
		error,
		mode: options.canEdit !== true || options.readOnly ? ('view' as const) : ('edit' as const),
	};
}
export type { SpreadsheetToolbarProps } from 'xlsx-web-component';

/** Optional controls using exactly the callbacks exposed to custom host controls. */
export function Toolbar(props: SpreadsheetToolbarProps) {
	return createElement(
		'div',
		{ role: 'toolbar', ref: (host: HTMLDivElement | null) => prepareBlock(host, 'toolbar') },
		...toolbarModel(props).map((item) =>
			item.kind === 'select'
				? createElement(
						'select',
						{
							key: item.id,
							'aria-label': item.label,
							disabled: item.disabled,
							value: item.value,
							onChange: (event: import('react').ChangeEvent<HTMLSelectElement>) => {
								void item.change(event.currentTarget.value);
							},
						},
						...item.options.map((o) =>
							createElement('option', { key: o.value, value: o.value }, o.label),
						),
					)
				: createElement(
						'button',
						{
							key: item.id,
							type: 'button',
							disabled: item.disabled,
							'aria-pressed': item.pressed === null ? 'mixed' : item.pressed,
							onMouseDown: (event: import('react').MouseEvent) => event.preventDefault(),
							onClick: () => {
								void item.run();
							},
						},
						item.label,
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
		{ ref: (host: HTMLDivElement | null) => prepareBlock(host, 'formula-bar') },
		createElement('input', {
			key: `name:${props.selection}`,
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
			key: `formula:${props.activeCell}`,
			ref: input,
			'aria-label': props.labels['Formula Bar'],
			defaultValue: props.formula,
			readOnly: !props.canEdit,
			'aria-invalid': invalid || props.editError !== null,
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
		props.editError === null
			? null
			: createElement('span', { key: 'edit-error', role: 'alert' }, props.editError),
	);
}
export function SheetTabs(props: SpreadsheetToolbarProps) {
	return createElement(
		'div',
		{
			role: 'tablist',
			'aria-label': props.labels['Sheet'],
			ref: (host: HTMLDivElement | null) => prepareBlock(host, 'tabs'),
		},
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
