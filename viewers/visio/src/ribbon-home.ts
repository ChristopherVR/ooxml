import { textGroups } from './ribbon-home-format.js';
import { command, commandRow, group, menu, stack, type CommandSpec } from './ribbon-parts.js';

const TEXT = 'Needs core text formatting edits.';
const STYLE = 'Needs core fill, line and effect edits.';
const ARRANGE = 'Needs core alignment, position, z-order and grouping edits.';
const CONNECT = 'Needs core connector and glue edits.';
const SHAPES = 'Needs core shape creation beyond rectangles.';
const SELECT = 'Needs multi-shape selection.';
const small = (spec: CommandSpec): CommandSpec => ({ size: 'small', ...spec });
const icon = (spec: CommandSpec): CommandSpec => ({ size: 'icon', ...spec });

/**
 * Microsoft Visio's Home tab, group for group: Clipboard, Font, Paragraph, Tools, Shape
 * Styles, Arrange and Editing. Commands the core cannot perform yet are shown disabled with the
 * missing capability in their tooltip. On phones every group folds into the Tools disclosure.
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
				small({ id: 'connector', label: 'Connector', icon: 'connector', unsupported: CONNECT }),
			),
			command(doc, small({ id: 'text-tool', label: 'Text', icon: 'textBox', unsupported: SHAPES })),
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
					{ id: 'ellipse', label: 'Ellipse', unsupported: SHAPES },
					{ id: 'line-tool', label: 'Line', icon: 'line', unsupported: SHAPES },
					{ id: 'freeform', label: 'Freeform', unsupported: SHAPES },
					{ id: 'arc', label: 'Arc', unsupported: SHAPES },
					{ id: 'pencil', label: 'Pencil', icon: 'pencil', unsupported: SHAPES },
				],
			}),
			command(
				doc,
				icon({
					id: 'connection-point',
					label: 'Connection Point',
					icon: 'visioConnectionPoint',
					unsupported: CONNECT,
				}),
			),
			command(
				doc,
				icon({ id: 'text-block', label: 'Text Block', icon: 'visioTextBlock', unsupported: TEXT }),
			),
		]),
	]);
	const styles = group(
		doc,
		'Shape Styles',
		[
			menu(doc, {
				id: 'quick-styles',
				label: 'Quick Styles',
				icon: 'quickStyles',
				unsupported: STYLE,
				items: [{ id: 'quick-style', label: 'Theme styles', unsupported: STYLE }],
			}),
			stack(doc, [
				menu(doc, {
					id: 'fill',
					label: 'Fill',
					icon: 'fill',
					size: 'small',
					unsupported: STYLE,
					items: [{ id: 'fill-options', label: 'Fill Options...', unsupported: STYLE }],
				}),
				menu(doc, {
					id: 'line',
					label: 'Line',
					icon: 'line',
					size: 'small',
					unsupported: STYLE,
					items: [{ id: 'line-options', label: 'Line Options...', unsupported: STYLE }],
				}),
				menu(doc, {
					id: 'effects',
					label: 'Effects',
					icon: 'effects',
					size: 'small',
					unsupported: STYLE,
					items: [{ id: 'shadow', label: 'Shadow', unsupported: STYLE }],
				}),
			]),
		],
		{ launcher: STYLE },
	);
	const arrange = group(doc, 'Arrange', [
		menu(doc, {
			id: 'align',
			label: 'Align',
			icon: 'alignObjects',
			unsupported: ARRANGE,
			items: [
				'Align Left',
				'Align Center',
				'Align Right',
				'Align Top',
				'Align Middle',
				'Align Bottom',
			].map((label) => ({
				id: label.toLowerCase().replace(' ', '-shapes-'),
				label,
				unsupported: ARRANGE,
			})),
		}),
		menu(doc, {
			id: 'position',
			label: 'Position',
			icon: 'position',
			unsupported: ARRANGE,
			items: [
				{ id: 'auto-align', label: 'Auto Align & Space', unsupported: ARRANGE },
				{ id: 'rotate', label: 'Rotate Shapes', unsupported: ARRANGE },
			],
		}),
		stack(doc, [
			menu(doc, {
				id: 'bring-to-front',
				label: 'Bring to Front',
				icon: 'bringToFront',
				size: 'small',
				split: true,
				unsupported: ARRANGE,
				items: [{ id: 'bring-forward', label: 'Bring Forward', unsupported: ARRANGE }],
			}),
			menu(doc, {
				id: 'send-to-back',
				label: 'Send to Back',
				icon: 'sendToBack',
				size: 'small',
				split: true,
				unsupported: ARRANGE,
				items: [{ id: 'send-backward', label: 'Send Backward', unsupported: ARRANGE }],
			}),
			menu(doc, {
				id: 'group',
				label: 'Group',
				icon: 'group',
				size: 'small',
				unsupported: ARRANGE,
				items: [{ id: 'ungroup', label: 'Ungroup', unsupported: ARRANGE }],
			}),
		]),
	]);
	const editing = group(doc, 'Editing', [
		menu(doc, {
			id: 'change-shape',
			label: 'Change Shape',
			icon: 'visioChangeShape',
			unsupported: 'Needs core master replacement.',
			items: [
				{ id: 'change-shape-item', label: 'Shapes', unsupported: 'Needs core master replacement.' },
			],
		}),
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
					{ id: 'replace', label: 'Replace...', unsupported: TEXT },
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
						action: { type: 'reveal', panel: 'layers' },
					},
					{
						id: 'assign-layer',
						label: 'Assign to Layer...',
						unsupported: 'Needs core layer assignment edits.',
					},
				],
			}),
			menu(doc, {
				id: 'select',
				label: 'Select',
				icon: 'pointer',
				size: 'small',
				items: [
					{ id: 'select-all', label: 'Select All', unsupported: SELECT },
					{ id: 'select-by-type', label: 'Select by Type...', unsupported: SELECT },
				],
			}),
		]),
	]);
	const tools0 = doc.createElement('details');
	tools0.className = 'ribbon-tools';
	tools0.open = true;
	const summary = doc.createElement('summary');
	const caret = doc.createElement('span');
	caret.setAttribute('aria-hidden', 'true');
	caret.textContent = '⌄';
	summary.append('Tools ', caret);
	const content = doc.createElement('div');
	content.className = 'tools-content';
	content.append(clipboard, font, paragraph, tools, styles, arrange, editing);
	tools0.append(summary, content);
	panel.append(commandRow(doc, 'Home commands', [tools0]));
}
