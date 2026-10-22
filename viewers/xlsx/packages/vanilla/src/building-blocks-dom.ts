import { prepareBlock } from '../../bindings/src/block-theme';
import {
	selectReference,
	toolbarModel,
	type SpreadsheetToolbarProps,
	type ToolbarItemModel,
} from '../../bindings/src/building-blocks';

/** Updates a host's controls from the latest props, patching elements so focus and drafts survive. */
export type BlockRenderer = (props: SpreadsheetToolbarProps | null) => void;

const optionsKey = (item: ToolbarItemModel): string =>
	item.kind === 'select' ? item.options.map((o) => `${o.value}\n${o.label}`).join('\n') : '';

export function toolbarRenderer(host: HTMLElement): BlockRenderer {
	prepareBlock(host, 'toolbar');
	host.setAttribute('role', 'toolbar');
	let controls: HTMLElement[] | undefined;
	let model: ToolbarItemModel[] = [];
	const keys: string[] = [];
	return (props) => {
		if (!props) {
			host.replaceChildren();
			controls = undefined;
			return;
		}
		model = toolbarModel(props);
		controls ??= model.map((item, index) => {
			if (item.kind === 'select') {
				const select = host.ownerDocument.createElement('select');
				select.onchange = () => {
					const current = model[index];
					if (current?.kind === 'select') void current.change(select.value);
				};
				return select;
			}
			const button = host.ownerDocument.createElement('button');
			button.type = 'button';
			// Keep focus (and the grid selection) where it was.
			button.onmousedown = (event) => event.preventDefault();
			button.onclick = () => {
				const current = model[index];
				if (current?.kind === 'button') void current.run();
			};
			return button;
		});
		if (host.childElementCount !== controls.length) host.replaceChildren(...controls);
		model.forEach((item, index) => {
			const control = controls![index]!;
			if (item.kind === 'select' && control instanceof HTMLSelectElement) {
				control.setAttribute('aria-label', item.label);
				control.disabled = item.disabled;
				const key = optionsKey(item);
				if (keys[index] !== key) {
					keys[index] = key;
					control.replaceChildren(
						...item.options.map((o) => {
							const option = host.ownerDocument.createElement('option');
							option.value = o.value;
							option.textContent = o.label;
							return option;
						}),
					);
				}
				control.value = item.value;
			} else if (item.kind === 'button') {
				(control as HTMLButtonElement).disabled = item.disabled;
				control.textContent = item.label;
				control.setAttribute(
					'aria-pressed',
					item.pressed === null ? 'mixed' : String(item.pressed),
				);
			}
		});
	};
}

export function formulaBarRenderer(host: HTMLElement): BlockRenderer {
	prepareBlock(host, 'formula-bar');
	const doc = host.ownerDocument;
	const name = doc.createElement('input');
	const formula = doc.createElement('input');
	const alert = doc.createElement('span');
	alert.setAttribute('role', 'alert');
	let latest: SpreadsheetToolbarProps | null = null;
	let invalidDraft = false;
	let lastSelection: string | undefined;
	const markFormula = () =>
		formula.setAttribute(
			'aria-invalid',
			String(invalidDraft || (latest?.editError ?? null) !== null),
		);
	name.oninput = () => name.setAttribute('aria-invalid', 'false');
	name.onkeydown = (event) => {
		if (event.key !== 'Enter' || !latest) return;
		event.preventDefault();
		name.setAttribute('aria-invalid', String(!selectReference(latest, name.value)));
	};
	formula.oninput = () => {
		invalidDraft = false;
		markFormula();
		latest?.onFormulaInput(formula.value, formula.selectionEnd ?? undefined);
	};
	formula.onkeydown = (event) => {
		if (!latest) return;
		if (event.key === 'Enter') {
			event.preventDefault();
			invalidDraft = !latest.onFormulaChange(formula.value);
		} else if (event.key === 'Escape') {
			event.preventDefault();
			latest.onCancelEdit();
			invalidDraft = false;
		} else return;
		markFormula();
	};
	host.append(name, formula);
	return (props) => {
		latest = props;
		if (!props) return;
		if (props.selection !== lastSelection) {
			name.value = props.selection;
			lastSelection = props.selection;
		}
		name.setAttribute('aria-label', props.labels['Name Box'] ?? 'Name Box');
		if (formula.value !== props.formula) formula.value = props.formula;
		formula.readOnly = !props.canEdit;
		formula.setAttribute('aria-label', props.labels['Formula Bar'] ?? 'Formula Bar');
		markFormula();
		alert.textContent = props.editError ?? '';
		if (props.editError === null) alert.remove();
		else if (!alert.isConnected) host.append(alert);
	};
}

export function sheetTabsRenderer(host: HTMLElement): BlockRenderer {
	prepareBlock(host, 'tabs');
	host.setAttribute('role', 'tablist');
	const buttons = new Map<number, HTMLButtonElement>();
	let latest: SpreadsheetToolbarProps | null = null;
	return (props) => {
		latest = props;
		if (!props) {
			host.replaceChildren();
			buttons.clear();
			return;
		}
		host.setAttribute('aria-label', props.labels['Sheet'] ?? 'Sheet');
		const visible = props.sheets.filter((sheet) => sheet.state === 'visible');
		for (const [index, button] of buttons)
			if (!visible.some((sheet) => sheet.index === index)) {
				button.remove();
				buttons.delete(index);
			}
		visible.forEach((sheet, position) => {
			let button = buttons.get(sheet.index);
			if (!button) {
				button = host.ownerDocument.createElement('button');
				button.type = 'button';
				button.setAttribute('role', 'tab');
				button.onclick = () => latest?.onSheetChange(sheet.index);
				buttons.set(sheet.index, button);
			}
			button.setAttribute('aria-selected', String(sheet.index === props.activeSheet));
			button.textContent = sheet.name;
			if (host.children[position] !== button)
				host.insertBefore(button, host.children[position] ?? null);
		});
	};
}
