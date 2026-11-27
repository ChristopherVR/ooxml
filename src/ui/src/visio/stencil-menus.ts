import type { OfficeUiContextMenu } from '../menu/context-menu';
import type { StencilAction } from './stencil-authoring';
import { findMaster, type Stencil } from './stencil-catalog';
import { stencilDrawing } from './stencil-drawing';
import { MY_SHAPES_ID, stencilErrorMessage, type StencilFiles } from './stencil-files';
import { menuItem } from './shapes-master-list';
import { loadStencilStrings, translateStencilText } from './stencil-locale';
import type { ViewerController } from './controller';

export type StencilMenu = HTMLElement &
	Partial<Pick<OfficeUiContextMenu, 'openAt' | 'open' | 'close'>>;

export function bindStencilMenus(
	pane: HTMLElement,
	files: StencilFiles,
	masterMenu: StencilMenu,
	options: {
		quickIds: (stencil: Stencil) => readonly string[];
		changeQuick: (stencil: Stencil, id: string) => void;
		error: (message: string) => void;
		update: (id: string) => void;
		controller: () => ViewerController | undefined;
	},
) {
	const doc = pane.ownerDocument;
	let menuMaster: string | undefined;
	let menuSection: string | undefined;
	let pending = Promise.resolve();
	const editing = new Set<string>();
	const t = (text: string) => translateStencilText(pane, text);
	const drawingSource = () => {
		const controller = options.controller();
		return controller && stencilDrawing(controller);
	};
	const run = (action: StencilAction) => {
		pending = pending
			.then(async () => {
				options.error('');
				const [{ runStencilAction }] = await Promise.all([
					import('./stencil-authoring'),
					loadStencilStrings(),
				]);
				await runStencilAction(pane, files, action, t, (id) => {
					if (action.type === 'new' || action.type === 'collect') editing.add(id);
					options.update(id);
				});
			})
			.catch((error: unknown) => options.error(stencilErrorMessage(error, t)));
	};
	const stencilMenu = doc.createElement('office-ui-context-menu') as StencilMenu;
	stencilMenu.dataset.shapesMenu = 'stencil';
	stencilMenu.setAttribute('label', t('Stencil actions'));
	const disabled = (command: string, label: string, reason: string | undefined) => {
		const item = menuItem(doc, command, t(label));
		if (reason) {
			item.setAttribute('disabled', '');
			item.setAttribute('description', t(reason));
		}
		return item;
	};

	const openMasterMenu = (button: HTMLElement, x: number, y: number) => {
		const id = button.dataset.master;
		if (!id || !masterMenu.openAt) return false;
		const section = button.closest<HTMLElement>('[data-stencil]')?.dataset.stencil;
		const current = section ? files.get(section) : undefined;
		const file = files.fileMaster(id);
		const found = findMaster(id);
		const source = drawingSource();
		const inDrawing = source?.documentMaster(id);
		if (!file && !found && !inDrawing) return false;
		const items: HTMLElement[] = [];
		if (found)
			items.push(
				menuItem(
					doc,
					'quick-shapes',
					t(
						options.quickIds(found.stencil).includes(id)
							? 'Remove from Quick Shapes'
							: 'Add to Quick Shapes',
					),
				),
			);
		if (section !== MY_SHAPES_ID)
			items.push(
				disabled(
					'my-shapes',
					'Add to My Shapes',
					inDrawing && !source?.sourceAvailable()
						? 'Open the source drawing to copy this master.'
						: undefined,
				),
			);
		if (inDrawing && !file && !found)
			items.push(disabled('edit-document', 'Edit Master', source?.editMasterReason(id)));
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
					section === MY_SHAPES_ID ? 'Remove from My Shapes' : 'Delete Master',
					reason,
				),
			);
		}
		if (!items.length) return false;
		masterMenu.replaceChildren(...items);
		masterMenu.openAt(x, y, button);
		if (!masterMenu.open) return false;
		menuMaster = id;
		menuSection = section;
		return true;
	};
	pane.addEventListener('keydown', (event) => {
		const target = event.target;
		if (!(target instanceof Element)) return;
		if (
			(event.ctrlKey || event.metaKey) &&
			!event.shiftKey &&
			!event.altKey &&
			event.key.toLowerCase() === 'z'
		) {
			const id = target.closest<HTMLElement>('[data-stencil]')?.dataset.stencil;
			if (id && files.get(id)?.editable) {
				event.preventDefault();
				event.stopPropagation();
				if (files.canUndo(id)) run({ type: 'undo', stencil: id });
			}
		}
	});
	stencilMenu.addEventListener('office-command', (event) => {
		event.stopPropagation();
		const command = (event as CustomEvent<{ command: string }>).detail.command;
		if (!menuSection) return;
		if (command === 'save-stencil-file') run({ type: 'save', stencil: menuSection });
		if (command === 'undo-stencil') run({ type: 'undo', stencil: menuSection });
		if (command === 'edit-stencil') {
			const wasEditing = editing.has(menuSection);
			if (wasEditing) editing.delete(menuSection);
			else editing.add(menuSection);
			stencilMenu
				.querySelector('[command="edit-stencil"]')
				?.setAttribute('checked', String(!wasEditing));
		}
		if (command === 'add-selection' && editing.has(menuSection)) {
			const drawing = drawingSource();
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
			const drawing = drawingSource();
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
		if (command === 'edit-document') {
			const drawing = drawingSource();
			const source = drawing?.documentMaster(master);
			const shape = source?.shapes[0];
			if (!drawing || !source || !shape || drawing.editMasterReason(master)) return;
			pending = pending
				.then(async () => {
					options.error('');
					const [{ askMasterFormat }] = await Promise.all([
						import('./stencil-master-dialog'),
						loadStencilStrings(),
					]);
					const format = await askMasterFormat(
						pane,
						shape,
						t,
						'Shapes that inherit from this master change with it.',
					);
					if (format && Object.keys(format).length) await drawing.editMaster(master, format);
				})
				.catch((error: unknown) => options.error(stencilErrorMessage(error, t)));
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
		openMasterMenu,
		create: () => run({ type: 'new' }),
		close: (id: string) => editing.delete(id),
		openSection(button: HTMLElement, id: string) {
			menuSection = id;
			const drawing = drawingSource();
			const reason = editing.has(id)
				? (drawing?.selectionReason() ??
					(drawing ? undefined : 'Select one editable 2D shape in a drawing.'))
				: 'Turn on Edit Stencil to change its masters.';
			stencilMenu.setAttribute('label', t('Stencil actions'));
			stencilMenu.replaceChildren(
				menuItem(doc, 'edit-stencil', t('Edit Stencil'), editing.has(id)),
				disabled('add-selection', 'Add Selected Shape', reason),
				disabled(
					'undo-stencil',
					'Undo',
					files.canUndo(id) ? undefined : 'No stencil edit to undo.',
				),
				menuItem(doc, 'save-stencil-file', t('Save Stencil')),
			);
			const box = button.getBoundingClientRect();
			stencilMenu.openAt?.(box.left, box.bottom, button);
		},
	};
}
