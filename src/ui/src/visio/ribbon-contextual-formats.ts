import { arrangeGroup, shapeStylesGroup } from './ribbon-arrange';
import { command, commandRow, group, menu, stack, type CommandSpec } from './ribbon-parts';
import { sizeFields } from './ribbon-size-fields';
import type { VisioPictureAction, VisioContainerOptionAction } from './ribbon-action';

const COMPRESSION =
	'Compressing a picture re-encodes its embedded image, which the editor does not do yet.';
const AUTO_BALANCE =
	'Auto Balance needs the picture histogram, which the editor does not compute yet.';

const sizePosition = (doc: Document) =>
	command(doc, {
		id: 'size-position',
		label: 'Size & Position',
		icon: 'sizePosition',
		action: { type: 'sizePosition' },
		checked: false,
	});
const sizeGroup = (doc: Document, extra: readonly HTMLElement[] = []) =>
	group(doc, 'Size', [...extra, sizeFields(doc), sizePosition(doc)]);

/** A command shown, disabled, with the reason it is missing. */
const missing = (doc: Document, id: string, label: string, icon: string, reason: string) =>
	menu(doc, { id, label, icon, size: 'small', unsupported: reason, items: [] });

/** Brightness and Contrast presets: Visio's -40% to +40%, stored as 0 to 1 around 0.5. */
const PRESETS: readonly [string, number][] = [
	['-40%', 0.1],
	['-20%', 0.3],
	['0%', 0.5],
	['+20%', 0.7],
	['+40%', 0.9],
];
const adjustMenu = (doc: Document, key: 'brightness' | 'contrast', label: string, icon: string) =>
	menu(doc, {
		id: `picture-${key}`,
		label,
		icon,
		size: 'small',
		items: PRESETS.map(([text, value]): CommandSpec => ({
			id: `picture-${key}-${value}`,
			label: text,
			action:
				key === 'brightness'
					? { type: 'picture-adjust', brightness: value }
					: { type: 'picture-adjust', contrast: value },
		})),
	});

const CROP_ITEMS: readonly [string, VisioPictureAction][] = [
	['Crop to Square', { type: 'picture-crop', crop: 'square' }],
	['Reset Crop', { type: 'picture-crop', crop: 'reset' }],
];

/** Container Format margins: none, the editor's narrow and normal margins and the wide one. */
export const CONTAINER_MARGINS: readonly [string, number][] = [
	['None', 0],
	['Narrow', 0.0625],
	['Normal', 3 / 25.4],
	['Wide', 0.25],
];
/** Automatic Resize: the three msvSDContainerResize modes, as Visio names them. */
const RESIZE: readonly [string, 0 | 1 | 2][] = [
	['No automatic resize', 0],
	['Expand as needed', 1],
	['Always fit to contents', 2],
];
const containerOption = (
	id: string,
	label: string,
	action: VisioContainerOptionAction,
): CommandSpec => ({ id, label, action });

/**
 * Picture Format: Adjust (Brightness and Contrast presets, Crop), Picture Styles (Line and
 * Effects, the Home controls built again), Arrange and Size. Compression and Auto Balance stay
 * listed disabled with the reason.
 */
export function buildPictureFormatPanel(doc: Document, panel: HTMLElement): void {
	panel.dataset.contextual = '';
	panel.append(
		commandRow(doc, 'Picture Format commands', [
			group(doc, 'Adjust', [
				stack(doc, [
					adjustMenu(doc, 'brightness', 'Brightness', 'effects'),
					adjustMenu(doc, 'contrast', 'Contrast', 'effects'),
					missing(doc, 'picture-auto-balance', 'Auto Balance', 'effects', AUTO_BALANCE),
				]),
				stack(doc, [
					missing(doc, 'picture-compression', 'Compression', 'visioPicture', COMPRESSION),
				]),
			]),
			shapeStylesGroup(doc, { label: 'Picture Styles', parts: ['line', 'effects'] }),
			arrangeGroup(doc),
			sizeGroup(doc, [
				menu(doc, {
					id: 'picture-crop',
					label: 'Crop',
					icon: 'visioPicture',
					size: 'small',
					items: CROP_ITEMS.map(([label, action], index): CommandSpec => ({
						id: `picture-crop-${index ? 'reset' : 'square'}`,
						label,
						action,
					})),
				}),
			]),
		]),
	);
}

/**
 * Container Format: Size (Margins, Fit to Contents and Automatic Resize, then the Width and
 * Height boxes), Container Styles (Fill, Line, Effects), Membership (Lock Container) and Arrange.
 */
export function buildContainerFormatPanel(doc: Document, panel: HTMLElement): void {
	panel.dataset.contextual = '';
	panel.append(
		commandRow(doc, 'Container Format commands', [
			group(doc, 'Size', [
				stack(doc, [
					menu(doc, {
						id: 'container-margins',
						label: 'Margins',
						icon: 'grid',
						size: 'small',
						items: CONTAINER_MARGINS.map(([label, margin], index) =>
							containerOption(`container-margin-${index}`, label, {
								type: 'container-options',
								margin,
							}),
						),
					}),
					command(doc, {
						id: 'container-fit',
						label: 'Fit to Contents',
						icon: 'fit',
						size: 'small',
						action: { type: 'container', command: 'fit' },
					}),
					menu(doc, {
						id: 'container-resize',
						label: 'Automatic Resize',
						icon: 'grid',
						size: 'small',
						items: RESIZE.map(([label, resize]) =>
							containerOption(`container-resize-${resize}`, label, {
								type: 'container-options',
								resize,
							}),
						),
					}),
				]),
				sizeFields(doc),
				sizePosition(doc),
			]),
			shapeStylesGroup(doc, {
				label: 'Container Styles',
				parts: ['fill', 'line', 'effects'],
			}),
			group(doc, 'Membership', [
				command(doc, {
					id: 'container-lock',
					label: 'Lock Container',
					icon: 'lock',
					size: 'small',
					action: { type: 'container-options', locked: 'toggle' },
				}),
				command(doc, {
					id: 'container-select-contents',
					label: 'Select Contents',
					icon: 'select',
					size: 'small',
					action: { type: 'container', command: 'select-contents' },
				}),
				command(doc, {
					id: 'container-disband',
					label: 'Disband Container',
					icon: 'delete',
					size: 'small',
					action: { type: 'container', command: 'disband' },
				}),
			]),
			arrangeGroup(doc),
		]),
	);
}

/**
 * Chart Design: the chart editor's entry points (the same one a double-click opens), Arrange and
 * Size. Chart commands are enabled while the chart is selected on an editable drawing.
 */
export function buildChartDesignPanel(doc: Document, panel: HTMLElement): void {
	panel.dataset.contextual = '';
	const chart = (
		id: string,
		label: string,
		icon: string,
		item: 'chart-edit' | 'chart-data' | 'chart-format',
	) => command(doc, { id, label, icon, action: { type: 'insert', item } });
	panel.append(
		commandRow(doc, 'Chart Design commands', [
			group(doc, 'Chart', [
				chart('chart-tab-edit', 'Edit Chart', 'visioChart', 'chart-edit'),
				chart('chart-tab-data', 'Edit Data', 'grid', 'chart-data'),
				chart('chart-tab-format', 'Format Chart', 'fill', 'chart-format'),
			]),
			arrangeGroup(doc),
			sizeGroup(doc),
		]),
	);
}
