import type {
	CloudFileProvider,
	CloudOpenedFile,
	OfficeConnectedServicesLabels,
	OfficeOpenPlacesLabels,
} from 'ooxml-ui/pptx';
import React, { useEffect, useLayoutEffect, useRef } from 'react';

type Host = HTMLElement & { value: string; checked: boolean; disabled: boolean };
type NativeEventHandler = (event: Event) => void;
type CommonProps = Omit<
	React.HTMLAttributes<HTMLElement>,
	'onChange' | 'onInput' | 'onFocus' | 'onBlur' | 'onKeyDown'
>;

function useHostEvent(
	ref: React.RefObject<Host | null>,
	type: string,
	handler?: NativeEventHandler,
): void {
	useEffect(() => {
		const host = ref.current;
		if (!host || !handler) {
			return;
		}
		host.addEventListener(type, handler);
		return () => host.removeEventListener(type, handler);
	}, [ref, type, handler]);
}

/** React 18 needs native listeners for custom element events. */
export function WebSearch({
	value,
	disabled = false,
	onInput,
	onChange,
	onFocus,
	onBlur,
	onKeyDown,
	...rest
}: CommonProps & {
	value: string;
	disabled?: boolean;
	placeholder?: string;
	variant?: 'titlebar';
	onInput?: NativeEventHandler;
	onChange?: NativeEventHandler;
	onFocus?: NativeEventHandler;
	onBlur?: NativeEventHandler;
	onKeyDown?: (event: KeyboardEvent) => void;
}): React.ReactElement {
	const ref = useRef<Host>(null);
	useLayoutEffect(() => {
		if (ref.current) {
			ref.current.value = value;
			ref.current.disabled = disabled;
		}
	}, [value, disabled]);
	useHostEvent(ref, 'input', onInput);
	useHostEvent(ref, 'change', onChange);
	useHostEvent(ref, 'focusin', onFocus);
	useHostEvent(ref, 'focusout', onBlur);
	useHostEvent(ref, 'keydown', onKeyDown as NativeEventHandler | undefined);
	return <pptx-ui-search ref={ref} value={value} disabled={disabled || undefined} {...rest} />;
}

/**
 * The places of File > Open (`<pptx-ui-open-places>`): its children are "This device", each
 * provider is a tab with the shared cloud file browser. A file opened there reaches `onOpen`.
 */
export function WebOpenPlaces({
	providers,
	accept,
	labels,
	locale,
	selected,
	onPlaceChange,
	onOpen,
	children,
	...rest
}: CommonProps & {
	providers: readonly CloudFileProvider[];
	accept: readonly string[];
	labels: OfficeOpenPlacesLabels;
	locale: string | undefined;
	selected: string | undefined;
	onPlaceChange: (place: string) => void;
	onOpen: (file: CloudOpenedFile) => void;
	children?: React.ReactNode;
}): React.ReactElement {
	const ref = useRef<HTMLElement & Record<string, unknown>>(null);
	useLayoutEffect(() => {
		const host = ref.current;
		if (host) {
			host.providers = providers;
			host.accept = accept;
			host.labels = labels;
			host.locale = locale;
			host.selected = selected ?? 'device';
		}
	}, [providers, accept, labels, locale, selected]);
	useEffect(() => {
		const host = ref.current;
		if (!host) {
			return;
		}
		const change = (event: Event): void =>
			onPlaceChange((event as CustomEvent<{ place: string }>).detail.place);
		const open = (event: Event): void =>
			onOpen((event as CustomEvent<{ file: CloudOpenedFile }>).detail.file);
		host.addEventListener('office-open-place-change', change);
		host.addEventListener('office-cloud-file-open', open);
		return () => {
			host.removeEventListener('office-open-place-change', change);
			host.removeEventListener('office-cloud-file-open', open);
		};
	}, [onPlaceChange, onOpen]);
	return (
		<pptx-ui-open-places ref={ref} {...rest}>
			{children}
		</pptx-ui-open-places>
	);
}

/**
 * File > Account's linked storage (`<pptx-ui-connected-services>`): one row per provider with its
 * own Connect and Disconnect. It renders nothing without providers.
 */
export function WebConnectedServices({
	providers,
	labels,
	...rest
}: CommonProps & {
	providers: readonly CloudFileProvider[];
	labels: OfficeConnectedServicesLabels;
}): React.ReactElement {
	const ref = useRef<HTMLElement & Record<string, unknown>>(null);
	useLayoutEffect(() => {
		const host = ref.current;
		if (host) {
			host.providers = providers;
			host.labels = labels;
		}
	}, [providers, labels]);
	return <pptx-ui-connected-services ref={ref} {...rest} />;
}

export function WebSelect({
	value,
	disabled = false,
	onChange,
	...rest
}: Omit<React.SelectHTMLAttributes<HTMLSelectElement>, 'value' | 'onChange'> & {
	value: string | number;
	variant?: 'ribbon-font';
	onChange?: React.ChangeEventHandler<HTMLSelectElement>;
}): React.ReactElement {
	const ref = useRef<Host>(null);
	useLayoutEffect(() => {
		if (ref.current) {
			ref.current.value = String(value);
			ref.current.disabled = disabled;
		}
	}, [value, disabled]);
	useHostEvent(
		ref,
		'change',
		onChange
			? (event) => onChange(event as unknown as React.ChangeEvent<HTMLSelectElement>)
			: undefined,
	);
	return (
		<pptx-ui-select ref={ref} value={String(value)} disabled={disabled || undefined} {...rest} />
	);
}

export function WebCheckbox({
	checked,
	disabled = false,
	onChange,
	value,
	...rest
}: Omit<React.InputHTMLAttributes<HTMLInputElement>, 'checked' | 'onChange'> & {
	checked: boolean;
	onChange?: React.ChangeEventHandler<HTMLInputElement>;
}): React.ReactElement {
	const ref = useRef<Host>(null);
	useLayoutEffect(() => {
		if (ref.current) {
			ref.current.checked = checked;
			ref.current.disabled = disabled;
		}
	}, [checked, disabled]);
	useHostEvent(
		ref,
		'change',
		onChange
			? (event) => onChange(event as unknown as React.ChangeEvent<HTMLInputElement>)
			: undefined,
	);
	return (
		<pptx-ui-checkbox
			ref={ref}
			checked={checked || undefined}
			disabled={disabled || undefined}
			value={value === undefined ? undefined : String(value)}
			{...rest}
		/>
	);
}

/** Shared radio: radios with one `name` form a roving-tabindex radiogroup. */
export function WebRadio({
	checked,
	disabled = false,
	onChange,
	value,
	name,
	...rest
}: Omit<React.InputHTMLAttributes<HTMLInputElement>, 'checked' | 'onChange' | 'type'> & {
	checked: boolean;
	onChange?: React.ChangeEventHandler<HTMLInputElement>;
}): React.ReactElement {
	const ref = useRef<Host>(null);
	useLayoutEffect(() => {
		if (ref.current) {
			ref.current.checked = checked;
			ref.current.disabled = disabled;
		}
	}, [checked, disabled]);
	useHostEvent(
		ref,
		'change',
		onChange
			? (event) => onChange(event as unknown as React.ChangeEvent<HTMLInputElement>)
			: undefined,
	);
	return (
		<pptx-ui-radio
			ref={ref}
			checked={checked || undefined}
			disabled={disabled || undefined}
			name={name}
			value={value === undefined ? undefined : String(value)}
			{...rest}
		/>
	);
}
