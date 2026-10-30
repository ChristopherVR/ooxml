import { textGroups } from './ribbon-home-format';
import { arrangeGroup, shapeStylesGroup } from './ribbon-arrange';
import { command, commandRow, group, menu, stack, type CommandSpec } from './ribbon-parts';
import { changeShapeGallery } from './ribbon-change-shape';

const small = (spec: CommandSpec): CommandSpec => ({ size: 'small', ...spec });
const icon = (spec: CommandSpec): CommandSpec => ({ size: 'icon', ...spec });

/**
 * Microsoft Visio's Home tab, group for group: Clipboard, Font, Paragraph, Tools, Shape
 * Styles, Arrange and Editing. Commands the core cannot perform yet are shown disabled with the
 * missing capability in their tooltip. On narrow windows the groups fold into dropdown buttons (ribbon-overflow.ts).
 */
export function buildHomePanel(doc: Document, panel: HTMLElement): void {
	const [clipboard, font, paragraph] = textGroups(doc);
	const tools = group(doc, 'Tools', [
		stack(doc, [
			command(
				doc,
				small({
					id: 'pointer',
					label: 'Pointer Tool',
					icon: 'pointer',
					action: { type: 'tool', tool: 'pointer' },
					keys: ['Control+1', 'Ctrl+1'],
					pressed: true,
				}),
			),
			command(
				doc,
				small({
					id: 'connector',
					label: 'Connector',
					icon: 'connector',
					action: { type: 'tool', tool: 'connector' },
					keys: ['Control+3', 'Ctrl+3'],
				}),
			),
			command(
				doc,
				small({
					id: 'text-tool',
					label: 'Text',
					icon: 'textBox',
					action: { type: 'tool', tool: 'text' },
					keys: ['Control+2', 'Ctrl+2'],
				}),
			),
		]),
		stack(doc, [
			menu(doc, {
				id: 'rectangle',
				label: 'Rectangle',
				icon: 'rectangle',
				size: 'icon',
				split: true,
				action: { type: 'tool', tool: 'rectangle' },
				keys: ['Control+8', 'Ctrl+8'],
				items: [
					{
						id: 'rectangle-item',
						label: 'Rectangle',
						icon: 'rectangle',
						action: { type: 'tool', tool: 'rectangle' },
						checked: false,
					},
					{
						id: 'ellipse',
						label: 'Ellipse',
						action: { type: 'tool', tool: 'ellipse' },
						keys: ['Control+9', 'Ctrl+9'],
						checked: false,
					},
					{
						id: 'line-tool',
						label: 'Line',
						icon: 'line',
						action: { type: 'tool', tool: 'line' },
						keys: ['Control+6', 'Ctrl+6'],
						checked: false,
					},
					{
						id: 'freeform',
						label: 'Freeform',
						action: { type: 'tool', tool: 'freeform' },
						keys: ['Control+5', 'Ctrl+5'],
						checked: false,
					},
					{
						id: 'arc',
						label: 'Arc',
						action: { type: 'tool', tool: 'arc' },
						keys: ['Control+7', 'Ctrl+7'],
						checked: false,
					},
					{
						id: 'pencil',
						label: 'Pencil',
						icon: 'pencil',
						action: { type: 'tool', tool: 'pencil' },
						keys: ['Control+4', 'Ctrl+4'],
						checked: false,
					},
				],
			}),
			command(
				doc,
				icon({
					id: 'connection-point',
					label: 'Connection Point',
					icon: 'visioConnectionPoint',
					action: { type: 'tool', tool: 'connection-point' },
					keys: ['Control+Shift+1', 'Ctrl+Shift+1'],
				}),
			),
			command(
				doc,
				icon({
					id: 'text-block',
					label: 'Text Block',
					icon: 'visioTextBlock',
					action: { type: 'text-feature', feature: 'text-block' },
					keys: ['Control+Shift+4', 'Ctrl+Shift+4'],
					pressed: false,
				}),
			),
		]),
	]);
	const styles = shapeStylesGroup(doc);
	const arrange = arrangeGroup(doc);
	const editing = group(doc, 'Editing', [
		changeShapeGallery(doc),
		stack(doc, [
			menu(doc, {
				id: 'find',
				label: 'Find',
				icon: 'search',
				size: 'small',
				items: [
					{
						id: 'find-item',
						label: 'Find...',
						icon: 'search',
						action: { type: 'search' },
						keys: ['Control+F', 'Ctrl+F'],
					},
					{
						id: 'replace',
						label: 'Replace...',
						action: { type: 'replace' },
						keys: ['Control+H', 'Ctrl+H'],
					},
				],
			}),
			menu(doc, {
				id: 'layers',
				label: 'Layers',
				icon: 'visioLayers',
				size: 'small',
				items: [
					{
						id: 'layer-properties',
						label: 'Layer Properties...',
						action: { type: 'layer-properties' },
					},
					{
						id: 'assign-layer',
						label: 'Assign to Layer...',
						action: { type: 'assign-layers' },
					},
				],
			}),
			menu(doc, {
				id: 'select',
				label: 'Select',
				icon: 'pointer',
				size: 'small',
				items: [
					{
						id: 'select-all',
						label: 'Select All',
						action: { type: 'selection', mode: 'all' },
						keys: ['Control+A', 'Ctrl+A'],
					},
					{
						id: 'clear-selection',
						label: 'Deselect All',
						action: { type: 'selection', mode: 'clear' },
					},
					{ id: 'select-by-type', label: 'Select by Type...', action: { type: 'select-by-type' } },
				],
			}),
		]),
	]);
	panel.append(
		commandRow(doc, 'Home commands', [clipboard, font, paragraph, tools, styles, arrange, editing]),
	);
}
