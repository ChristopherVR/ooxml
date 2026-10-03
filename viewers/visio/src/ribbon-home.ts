import { commandRow, group, hint, tool } from './ribbon-parts.js';
import { createSearchGroup } from './viewer-search.js';

/**
 * Visio's Home tab, limited to commands this build can honestly perform: Undo, Tools and
 * Editing. Clipboard, Font, Paragraph, Shape Styles and Arrange wait for core edit commands.
 * The compact (phone) layout folds every group but Undo into the Tools disclosure.
 */
export function buildHomePanel(doc: Document, panel: HTMLElement): void {
	const tools = doc.createElement('details');
	tools.className = 'ribbon-tools';
	tools.open = true;
	const summary = doc.createElement('summary');
	const caret = doc.createElement('span');
	caret.setAttribute('aria-hidden', 'true');
	caret.textContent = '⌄';
	summary.append('Tools ', caret);
	const content = doc.createElement('div');
	content.className = 'tools-content';
	content.append(
		group(doc, 'Tools', [
			tool(
				doc,
				'pointer',
				'Pointer Tool',
				'pointer',
				{ type: 'tool', tool: 'pointer' },
				{
					keys: ['Control+1', 'Ctrl+1'],
					pressed: true,
				},
			),
			tool(
				doc,
				'rectangle',
				'Rectangle',
				'rectangle',
				{ type: 'tool', tool: 'rectangle' },
				{
					keys: ['Control+8', 'Ctrl+8'],
					pressed: false,
				},
			),
			tool(
				doc,
				'edit',
				'Edit text',
				'pencil',
				{ type: 'reveal', panel: 'edit' },
				{
					keys: ['F2', 'F2'],
				},
			),
		]),
		createSearchGroup(doc),
		group(doc, 'Editing', [
			tool(doc, 'delete', 'Delete', 'trash', { type: 'delete' }, { keys: ['Delete', 'Del'] }),
			tool(doc, 'layers', 'Layers', 'visioLayers', { type: 'reveal', panel: 'layers' }),
			tool(doc, 'selection', 'Shape details', 'visioShapeData', {
				type: 'reveal',
				panel: 'selection',
			}),
		]),
	);
	tools.append(summary, content);
	panel.append(
		commandRow(doc, 'Home commands', [
			group(
				doc,
				'Undo',
				[
					tool(
						doc,
						'undo',
						'Undo',
						'undo',
						{ type: 'history', key: 'undo' },
						{
							keys: ['Control+Z', 'Ctrl+Z'],
						},
					),
					tool(
						doc,
						'redo',
						'Redo',
						'redo',
						{ type: 'history', key: 'redo' },
						{
							keys: ['Control+Y', 'Ctrl+Y'],
						},
					),
				],
				'undo-group',
			),
			tools,
		]),
		hint(doc, [
			'Rectangle: drag on the page to draw.',
			'Editing is experimental. Keep your original file.',
		]),
	);
}
