import { prepareBlock } from '../../bindings/src/block-theme';
import { createEffect, createSignal, onCleanup, onMount, untrack } from 'solid-js';
import {
	createBuildingBlockSession,
	emptyBuildingBlockState,
	toolbarModel,
	selectReference,
	type BuildingBlockOptions,
	type SpreadsheetToolbarProps,
} from '../../bindings/src/building-blocks';
export type {
	BuildingBlockOptions,
	SpreadsheetToolbarProps,
} from '../../bindings/src/building-blocks';
/** Call inside a Solid owner. Inputs and returned snapshots are accessors. */
export function createViewerBuildingBlocks(options: () => BuildingBlockOptions = () => ({})) {
	const [state, setState] = createSignal(emptyBuildingBlockState());
	const session = createBuildingBlockSession(setState);
	createEffect(() => session.update(options()));
	onCleanup(session.destroy);
	return {
		canvasProps: { mount: session.mount },
		handle: session.handle,
		toolbarProps: () => state().toolbarProps,
		loading: () => state().loading,
		error: () => state().error,
		mode: () => state().mode,
		dirty: () => state().dirty,
		selection: () => state().selection,
		collaboration: () => state().collaboration,
		editError: () => state().editError,
	};
}
export function SpreadsheetCanvas(props: {
	mount: (host: HTMLElement | null) => void;
	class?: string;
}) {
	if (typeof document === 'undefined') return undefined;
	const host = document.createElement('div');
	createEffect(() => {
		host.className = props.class ?? '';
	});
	const [mounted, setMounted] = createSignal(false);
	let current: ((host: HTMLElement | null) => void) | undefined;
	onMount(() => setMounted(true));
	createEffect(() => {
		if (!mounted()) return;
		const next = props.mount;
		if (current === next) return;
		untrack(() => {
			current?.(null);
			current = next;
			next(host);
		});
	});
	onCleanup(() => {
		current?.(null);
		current = undefined;
	});
	return host;
}
export function Toolbar(props: { toolbarProps: SpreadsheetToolbarProps }) {
	if (typeof document === 'undefined') return undefined;
	const host = document.createElement('div');
	prepareBlock(host, 'toolbar');
	host.setAttribute('role', 'toolbar');
	// The structure of the controls is fixed, so each element is created once and only updated.
	const model = () => toolbarModel(props.toolbarProps);
	const controls = untrack(model).map((item, index) => {
		if (item.kind === 'select') {
			const select = document.createElement('select');
			select.onchange = () => {
				const current = model()[index];
				if (current?.kind === 'select') void current.change(select.value);
			};
			return select;
		}
		const button = document.createElement('button');
		button.type = 'button';
		button.onmousedown = (e) => e.preventDefault();
		button.onclick = () => {
			const current = model()[index];
			if (current?.kind === 'button') void current.run();
		};
		return button;
	});
	host.append(...controls);
	createEffect(() => {
		model().forEach((item, index) => {
			const control = controls[index]!;
			if (item.kind === 'select' && control instanceof HTMLSelectElement) {
				control.setAttribute('aria-label', item.label);
				control.disabled = item.disabled;
				control.replaceChildren(
					...item.options.map((o) => {
						const option = document.createElement('option');
						option.value = o.value;
						option.textContent = o.label;
						return option;
					}),
				);
				control.value = item.value;
			} else if (item.kind === 'button' && control instanceof HTMLButtonElement) {
				control.disabled = item.disabled;
				control.textContent = item.label;
				control.setAttribute(
					'aria-pressed',
					item.pressed === null ? 'mixed' : String(item.pressed),
				);
			}
		});
	});
	return host;
}
export function FormulaBar(props: { toolbarProps: SpreadsheetToolbarProps }) {
	if (typeof document === 'undefined') return undefined;
	const host = document.createElement('div');
	prepareBlock(host, 'formula-bar');
	const name = document.createElement('input');
	const formula = document.createElement('input');
	const alert = document.createElement('span');
	alert.setAttribute('role', 'alert');
	host.append(name, formula);
	let invalidDraft = false;
	const markFormula = () =>
		formula.setAttribute(
			'aria-invalid',
			String(invalidDraft || props.toolbarProps.editError !== null),
		);
	let lastSelection: string | undefined;
	createEffect(() => {
		if (props.toolbarProps.selection !== lastSelection) {
			name.value = props.toolbarProps.selection;
			lastSelection = props.toolbarProps.selection;
		}
		name.setAttribute('aria-label', props.toolbarProps.labels['Name Box'] ?? 'Name Box');
	});
	createEffect(() => {
		const p = props.toolbarProps;
		if (formula.value !== p.formula) formula.value = p.formula;
		formula.readOnly = !p.canEdit;
		markFormula();
		alert.textContent = p.editError ?? '';
		if (p.editError === null) alert.remove();
		else if (!alert.isConnected) host.append(alert);
		formula.setAttribute('aria-label', p.labels['Formula Bar'] ?? 'Formula Bar');
	});
	name.oninput = () => name.setAttribute('aria-invalid', 'false');
	name.onkeydown = (e) => {
		if (e.key === 'Enter') {
			e.preventDefault();
			name.setAttribute('aria-invalid', String(!selectReference(props.toolbarProps, name.value)));
		}
	};
	formula.oninput = () => {
		invalidDraft = false;
		markFormula();
		props.toolbarProps.onFormulaInput(formula.value, formula.selectionEnd ?? undefined);
	};
	formula.onkeydown = (e) => {
		if (e.key === 'Enter') {
			e.preventDefault();
			invalidDraft = !props.toolbarProps.onFormulaChange(formula.value);
			markFormula();
		} else if (e.key === 'Escape') {
			e.preventDefault();
			props.toolbarProps.onCancelEdit();
			invalidDraft = false;
			markFormula();
		}
	};
	return host;
}
export function SheetTabs(props: { toolbarProps: SpreadsheetToolbarProps }) {
	if (typeof document === 'undefined') return undefined;
	const host = document.createElement('div');
	prepareBlock(host, 'tabs');
	host.setAttribute('role', 'tablist');
	const buttons = new Map<number, HTMLButtonElement>();
	createEffect(() => {
		host.setAttribute('aria-label', props.toolbarProps.labels['Sheet'] ?? 'Sheet');
		const visible = props.toolbarProps.sheets.filter((s) => s.state === 'visible');
		for (const [index, button] of buttons)
			if (!visible.some((s) => s.index === index)) {
				button.remove();
				buttons.delete(index);
			}
		visible.forEach((s, position) => {
			let button = buttons.get(s.index);
			if (!button) {
				button = document.createElement('button');
				button.type = 'button';
				button.setAttribute('role', 'tab');
				button.onclick = () => props.toolbarProps.onSheetChange(s.index);
				buttons.set(s.index, button);
			}
			button.setAttribute('aria-selected', String(s.index === props.toolbarProps.activeSheet));
			button.textContent = s.name;
			if (host.children[position] !== button)
				host.insertBefore(button, host.children[position] ?? null);
		});
	});
	return host;
}
