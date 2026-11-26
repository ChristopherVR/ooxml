import type { StencilAction } from './stencil-authoring';
import { findMaster, type Stencil } from './stencil-catalog';
import { stencilDrawing } from './stencil-drawing';
import { MY_SHAPES_ID, stencilErrorMessage, type StencilFiles } from './stencil-files';
import { menuItem } from './shapes-master-list';
import { ribbonLocale } from './ribbon-localize';
import { translateRibbonText } from './ribbon-locale';

type Menu = HTMLElement & { openAt?(x: number, y: number, opener?: HTMLElement): void };

export function bindStencilMenus(
	pane: HTMLElement,
	files: StencilFiles,
	masterMenu: Menu,
	options: {
		quickIds: (stencil: Stencil) => readonly string[];
		changeQuick: (stencil: Stencil, id: string) => void;
		error: (message: string) => void;
		update: (id: string) => void;
	},
) {
	const doc = pane.ownerDocument;
	let menuMaster: string | undefined;
	let menuSection: string | undefined;
	let pending = Promise.resolve();
	const editing = new Set<string>();
	const t = (text: string) => translateRibbonText(ribbonLocale(pane), text);
	const run = (action: StencilAction) => {
		pending = pending
			.then(async () => {
				options.error('');
				const { runStencilAction } = await import('./stencil-authoring');
				await runStencilAction(pane, files, action, t, (id) => {
					if (action.type === 'new' || action.type === 'collect') editing.add(id);
					options.update(id);
				});
			})
			.catch((error: unknown) => options.error(stencilErrorMessage(error, t)));
	};
	const stencilMenu = doc.createElement('office-ui-context-menu') as Menu;
	stencilMenu.dataset.shapesMenu = 'stencil';
	stencilMenu.setAttribute('label', t('Stencil actions'));
	const disabled = (command: string, label: string, reason: string | undefined) => {
		const item = menuItem(doc, command, t(label));
		if (reason) {
			item.setAttribute('disabled', '');
			item.title = t(reason);
		}
		return item;
	};

	const openMasterMenu = (button: HTMLElement, x: number, y: number) => {
		const id = button.dataset.master;
		if (!id) return false;
		menuMaster = id;
		menuSection = button.closest<HTMLElement>('[data-stencil]')?.dataset.stencil;
		const current = menuSection ? files.get(menuSection) : undefined;
		const file = files.fileMaster(id);
		const found = findMaster(id);
		const source = stencilDrawing(pane);
		const inDrawing = source?.documentMaster(id);
		if (!file && !found && !inDrawing) return false;
		const items: HTMLElement[] = [];
		if (found)
			items.push(
				menuItem(
					doc,
					'quick-shapes',
					options.quickIds(found.stencil).includes(id)
						? 'Remove from Quick Shapes'
						: 'Add to Quick Shapes',
				),
			);
		if (menuSection !== MY_SHAPES_ID)
			items.push(
				disabled(
					'my-shapes',
					'Add to My Shapes',
					inDrawing && !source?.sourceAvailable()
						? 'Open the source drawing to copy this master.'
						: undefined,
				),
			);
		if (file && current?.editable) {
			const active = editing.has(current.id);
			const reason = active ? undefined : 'Turn on Edit Stencil to change its masters.';
			items.push(disabled('rename', 'Rename Master', reason));
			const plain = file.master.rootCount === 1 && file.master.shapes[0]?.kind === 'shape';
			items.push(
				disabled(
					'edit',
					'Edit Master',
					reason ?? (plain ? undefined : 'Select a top-level shape with editable plain text.'),
				),
			);
			items.push(
				disabled(
					'delete',
					menuSection === MY_SHAPES_ID ? 'Remove from My Shapes' : 'Delete Master',
					reason,
				),
			);
		}
		masterMenu.replaceChildren(...items);
		masterMenu.openAt?.(x, y, button);
		return true;
	};
	pane.addEventListener('contextmenu', (event) => {
		const button = (event.target as Element).closest?.<HTMLElement>('[data-master]');
		if (button && openMasterMenu(button, event.clientX, event.clientY)) event.preventDefault();
	});
	pane.addEventListener('keydown', (event) => {
		if (event.key !== 'ContextMenu' && !(event.shiftKey && event.key === 'F10')) return;
		const button = (event.target as Element).closest?.<HTMLElement>('[data-master]');
		if (!button) return;
		const box = button.getBoundingClientRect();
		if (openMasterMenu(button, box.left, box.bottom)) event.preventDefault();
	});
	stencilMenu.addEventListener('office-command', (event) => {
		event.stopPropagation();
		const command = (event as CustomEvent<{ command: string }>).detail.command;
		if (!menuSection) return;
		if (command === 'save-stencil-file') run({ type: 'save', stencil: menuSection });
		if (command === 'edit-stencil') {
			if (editing.has(menuSection)) editing.delete(menuSection);
			else editing.add(menuSection);
		}
		if (command === 'add-selection' && editing.has(menuSection)) {
			const drawing = stencilDrawing(pane);
			if (drawing && !drawing.selectionReason())
				run({ type: command, stencil: menuSection, draft: drawing.selection });
		}
	});
	masterMenu.addEventListener('office-command', (event) => {
		event.stopPropagation();
		const command = (event as CustomEvent<{ command: string }>).detail.command;
		const master = menuMaster;
		if (!master) return;
		if (command === 'my-shapes') {
			const file = files.fileMaster(master);
			const found = findMaster(master);
			const drawing = stencilDrawing(pane);
			const source = drawing?.documentMaster(master);
			if (file)
				run({
					type: 'collect',
					draft: () => ({ type: 'import-master', stencil: file.bytes, id: file.master.id }),
				});
			else if (found) run({ type: 'collect', draft: () => ({ type: 'add-built-in', id: master }) });
			else if (source && drawing?.sourceAvailable())
				run({ type: 'collect', draft: () => drawing.captureMaster(source.id) });
			return;
		}
		if (
			(command === 'rename' || command === 'edit' || command === 'delete') &&
			menuSection &&
			editing.has(menuSection)
		) {
			run({ type: command, stencil: menuSection, master });
			return;
		}
		const found = findMaster(master);
		if (!found || command !== 'quick-shapes') return;
		options.changeQuick(found.stencil, master);
	});

	return {
		menu: stencilMenu,
		create: () => run({ type: 'new' }),
		close: (id: string) => editing.delete(id),
		openSection(button: HTMLElement, id: string) {
			menuSection = id;
			const drawing = stencilDrawing(pane);
			const reason = editing.has(id)
				? (drawing?.selectionReason() ??
					(drawing ? undefined : 'Select one editable 2D shape in a drawing.'))
				: 'Turn on Edit Stencil to change its masters.';
			stencilMenu.setAttribute('label', t('Stencil actions'));
			stencilMenu.replaceChildren(
				menuItem(doc, 'edit-stencil', t('Edit Stencil'), editing.has(id)),
				disabled('add-selection', 'Add Selected Shape', reason),
				menuItem(doc, 'save-stencil-file', t('Save Stencil')),
			);
			const box = button.getBoundingClientRect();
			stencilMenu.openAt?.(box.left, box.bottom, button);
		},
	};
}
