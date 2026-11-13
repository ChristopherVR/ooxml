// Host-owned context menu and workflow panels for the building-blocks demo (`?blocks=1&hostui=1`).
// Written once against `window.xlsxNativeBlocks`, so every framework mount gets the same UI.
import type { BuildingBlockOptions } from '../../../viewers/xlsx/packages/bindings/src/building-blocks';
import type { NativeHarness } from './native-building-blocks';

type Request = Parameters<NonNullable<BuildingBlockOptions['onContextMenu']>>[0];

/** `host` replaces the native menu; `native` keeps it and filters/extends it. */
export interface HostUiState {
	mode: 'host' | 'native';
	log: string[];
}
declare global {
	interface Window {
		xlsxHostUi: HostUiState;
	}
}

const blocks = (): NativeHarness => window.xlsxNativeBlocks;
const el = <K extends keyof HTMLElementTagNameMap>(
	tag: K,
	attrs: Record<string, string> = {},
	text = '',
): HTMLElementTagNameMap[K] => {
	const node = document.createElement(tag);
	for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, value);
	node.textContent = text;
	return node;
};

function closeMenu(): void {
	document.querySelector('#host-menu')?.remove();
}

/** The host's own menu: a list of buttons at the click position. */
function showHostMenu(request: Request): void {
	closeMenu();
	const menu = el('div', {
		id: 'host-menu',
		role: 'menu',
		'data-kind': request.menu,
		style: `position:fixed;z-index:50;left:${request.x}px;top:${request.y}px;background:#fff;color:#000;border:1px solid #888;max-height:60vh;overflow:auto`,
	});
	for (const item of request.items) {
		const button = el('button', { role: 'menuitem', 'data-id': item.id }, item.label);
		button.style.display = 'block';
		if (!item.enabled) button.setAttribute('aria-disabled', 'true');
		button.addEventListener('click', async () => {
			closeMenu();
			const ran = await request.execute(item.id);
			window.xlsxHostUi.log.push(`${request.menu}:${item.id}:${ran}`);
		});
		menu.append(button);
	}
	document.body.append(menu);
}

const props = () => blocks().snapshot().toolbarProps!;

function findPanel(): HTMLElement {
	const panel = el('section', { 'data-panel': 'find' });
	panel.append(el('h4', {}, 'Host Find'));
	const find = el('input', { 'aria-label': 'Host find what' });
	const replace = el('input', { 'aria-label': 'Host replace with' });
	const status = el('output', { 'aria-label': 'Host find status' });
	const query = () => ({ text: find.value });
	// Unchecked boxes pass no option, so the editor's `validateReplace` decides.
	const validate = el('input', { type: 'checkbox', 'aria-label': 'Host validate replace' });
	const skip = el('input', { type: 'checkbox', 'aria-label': 'Host skip validation' });
	let refused: string[] = [];
	const options = () => ({
		onRefused: (cells: Array<{ ref: string }>) => (refused = cells.map((cell) => cell.ref)),
		...(validate.checked ? { validate: true } : skip.checked ? { validate: false } : {}),
	});
	const buttons: Array<[string, () => string]> = [
		['Host find all', () => `${props().workflows.find.findAll(query()).length} found`],
		[
			'Host find next',
			() => {
				const match = props().workflows.find.findNext(query());
				return match ? `at ${match.row},${match.col}` : 'none';
			},
		],
		[
			'Host replace',
			() => `replaced ${props().workflows.find.replace(query(), replace.value, options())}`,
		],
		[
			'Host replace all',
			() => `replaced ${props().workflows.find.replaceAll(query(), replace.value, options())}`,
		],
	];
	panel.append(find, replace, validate, skip);
	for (const [label, run] of buttons) {
		const button = el('button', {}, label);
		button.addEventListener('click', () => {
			refused = [];
			const result = run();
			status.textContent = `${result}; refused=${refused.join(',')}; canReplace=${props().workflows.find.canReplace()}`;
		});
		panel.append(button);
	}
	panel.append(status);
	return panel;
}

function formatPanel(): HTMLElement {
	const panel = el('section', { 'data-panel': 'format' });
	panel.append(el('h4', {}, 'Host Format Cells'));
	const bold = el('input', { type: 'checkbox', 'aria-label': 'Host bold' });
	const numFmt = el('input', { 'aria-label': 'Host number format' });
	const status = el('output', { 'aria-label': 'Host format status' });
	const load = el('button', {}, 'Host load format');
	load.addEventListener('click', () => {
		const state = props().workflows.formatCells.getState();
		bold.checked = state.font.value.bold === true;
		bold.indeterminate = state.font.mixed.includes('bold');
		numFmt.value = state.numFmt ?? '';
		numFmt.placeholder = state.numFmt === null ? 'mixed' : '';
		status.textContent = `mixed=${state.font.mixed.join(',')};numFmt=${state.numFmt ?? 'mixed'}`;
	});
	const apply = el('button', {}, 'Host apply format');
	apply.addEventListener('click', () => {
		const ok = props().workflows.formatCells.apply({
			font: { bold: bold.checked },
			...(numFmt.value ? { numFmt: numFmt.value } : {}),
		});
		status.textContent = `applied=${ok}`;
	});
	panel.append(bold, numFmt, load, apply, status);
	return panel;
}

function functionPanel(): HTMLElement {
	const panel = el('section', { 'data-panel': 'function' });
	panel.append(el('h4', {}, 'Host Insert Function'));
	const text = el('input', { 'aria-label': 'Host function search' });
	const list = el('ul', { 'aria-label': 'Host function results' });
	const status = el('output', { 'aria-label': 'Host function status' });
	text.addEventListener('input', () => {
		list.replaceChildren();
		for (const doc of props().workflows.insertFunction.search({ text: text.value }).slice(0, 5)) {
			const item = el('li');
			const button = el('button', { 'data-function': doc.name }, `Host insert ${doc.name}`);
			// Keep the formula draft focused: the click must not move the caret.
			button.addEventListener('mousedown', (event) => event.preventDefault());
			button.addEventListener('click', () => {
				const result = props().workflows.insertFunction.insert(doc.name);
				status.textContent = result ? `${result.text}|${result.caret}` : 'refused';
			});
			item.append(button);
			list.append(item);
		}
	});
	panel.append(text, list, status);
	return panel;
}

/** The callbacks that make the demo host-owned; merged into every `update`. */
export function hostCallbacks(): Partial<BuildingBlockOptions> {
	window.xlsxHostUi = { mode: 'host', log: [] };
	return {
		onContextMenu: (request) => {
			if (window.xlsxHostUi.mode !== 'host') return false;
			showHostMenu(request);
			return true;
		},
		onContextMenuItems: (request) => {
			window.xlsxHostUi.log.push(`native:${request.menu}`);
			return {
				hide: request.items.some((item) => item.id === 'copy') ? ['copy'] : [],
				add: [
					{
						id: 'stamp',
						label: 'Host stamp',
						onSelect: (r) => void window.xlsxHostUi.log.push(`stamp:${r.menu}`),
					},
					{
						id: 'stamp-edit',
						label: 'Host stamp (edit)',
						requiresEdit: true,
						position: 'bottom',
						onSelect: (r) => void window.xlsxHostUi.log.push(`stamp-edit:${r.menu}`),
					},
				],
			};
		},
		onOpenFormatCells: () => {
			window.xlsxHostUi.log.push('open-format-cells');
		},
	};
}

/** Draws the three host panels; they read the live operations from the snapshot when used. */
export function mountHostPanels(): () => void {
	const root = el('div', {
		id: 'host-panels',
		style:
			'position:fixed;right:0;bottom:0;width:300px;max-height:45vh;overflow:auto;z-index:20;background:#fff;color:#000;font-size:12px',
	});
	root.append(findPanel(), formatPanel(), functionPanel());
	// The building-blocks canvas has the height of its parent; give the demo a real grid.
	const style = el('style', {}, 'xlsx-editor { display: block; height: 420px; }');
	document.head.append(style);
	document.body.append(root);
	const dismiss = (event: Event) => {
		if (!(event.target as Element).closest?.('#host-menu')) closeMenu();
	};
	// The tab strip block is the host's own buttons, so the host opens the tab menu itself.
	const onTabMenu = (event: MouseEvent) => {
		const tab = (event.target as Element).closest?.('[role="tablist"] [role="tab"]');
		if (!tab || tab.closest('xlsx-editor') || window.xlsxHostUi.mode !== 'host') return;
		const index = props().sheets.findIndex((sheet) => sheet.name === tab.textContent?.trim());
		if (index < 0) return;
		event.preventDefault();
		showHostMenu({ ...props().contextMenu('tab', index), x: event.clientX, y: event.clientY });
	};
	const onKey = (event: KeyboardEvent) => event.key === 'Escape' && closeMenu();
	document.addEventListener('keydown', onKey, true);
	document.addEventListener('contextmenu', onTabMenu);
	document.addEventListener('pointerdown', dismiss, true);
	return () => {
		document.removeEventListener('keydown', onKey, true);
		document.removeEventListener('contextmenu', onTabMenu);
		document.removeEventListener('pointerdown', dismiss, true);
		root.remove();
		style.remove();
		closeMenu();
	};
}
