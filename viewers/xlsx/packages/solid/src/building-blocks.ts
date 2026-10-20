import { createEffect, createSignal, onCleanup, onMount, untrack } from 'solid-js';
import {
	createBuildingBlockSession,
	emptyBuildingBlockState,
	toolbarButtons,
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
	host.setAttribute('role', 'toolbar');
	const buttons = Array.from({ length: 5 }, (_, index) => {
		const button = document.createElement('button');
		button.type = 'button';
		button.onmousedown = (e) => e.preventDefault();
		button.onclick = () => {
			void toolbarButtons(props.toolbarProps)[index]?.run();
		};
		return button;
	});
	host.append(...buttons);
	createEffect(() => {
		toolbarButtons(props.toolbarProps).forEach((b, index) => {
			const button = buttons[index]!;
			button.disabled = b.disabled;
			button.textContent = b.label;
			button.setAttribute('aria-pressed', b.pressed === null ? 'mixed' : String(b.pressed));
		});
	});
	return host;
}
export function FormulaBar(props: { toolbarProps: SpreadsheetToolbarProps }) {
	if (typeof document === 'undefined') return undefined;
	const host = document.createElement('div');
	const name = document.createElement('input');
	const formula = document.createElement('input');
	host.append(name, formula);
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
		formula.setAttribute('aria-invalid', 'false');
		props.toolbarProps.onFormulaInput(formula.value, formula.selectionEnd ?? undefined);
	};
	formula.onkeydown = (e) => {
		if (e.key === 'Enter') {
			e.preventDefault();
			formula.setAttribute(
				'aria-invalid',
				String(!props.toolbarProps.onFormulaChange(formula.value)),
			);
		} else if (e.key === 'Escape') {
			e.preventDefault();
			props.toolbarProps.onCancelEdit();
			formula.setAttribute('aria-invalid', 'false');
		}
	};
	return host;
}
export function SheetTabs(props: { toolbarProps: SpreadsheetToolbarProps }) {
	if (typeof document === 'undefined') return undefined;
	const host = document.createElement('div');
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
