import { arrangeGroup, shapeStylesGroup } from './ribbon-arrange';
import { command, commandRow, group, menu, stack } from './ribbon-parts';

const PICTURE =
	'Visio stores a picture as an embedded image that this editor does not recolour or recompress; replace it with an edited picture.';
const CROP = 'Cropping needs the picture clip cells, which the editor does not write yet.';
const MARGINS = 'The container margin (User.msvSDContainerMargin) cannot be edited yet.';
const RESIZE = 'The container resize mode (User.msvSDContainerResize) cannot be edited yet.';
const LOCK = 'Locking a container needs the shape protection cells, which are not editable yet.';

const sizePosition = (doc: Document) =>
	command(doc, {
		id: 'size-position',
		label: 'Size & Position',
		icon: 'sizePosition',
		action: { type: 'sizePosition' },
		checked: false,
	});
const sizeGroup = (doc: Document, extra: readonly HTMLElement[] = []) =>
	group(doc, 'Size', [...extra, sizePosition(doc)]);

/** A command shown, disabled, with the reason it is missing. */
const missing = (doc: Document, id: string, label: string, icon: string, reason: string) =>
	menu(doc, { id, label, icon, size: 'small', unsupported: reason, items: [] });

/**
 * Picture Format: Adjust (listed with the reason each is missing), Picture Styles (Line and
 * Effects, the Home controls built again), Arrange and Size. Crop is listed disabled.
 */
export function buildPictureFormatPanel(doc: Document, panel: HTMLElement): void {
	panel.dataset.contextual = '';
	panel.append(
		commandRow(doc, 'Picture Format commands', [
			group(doc, 'Adjust', [
				stack(doc, [
					missing(doc, 'picture-brightness', 'Brightness', 'effects', PICTURE),
					missing(doc, 'picture-contrast', 'Contrast', 'effects', PICTURE),
					missing(doc, 'picture-auto-balance', 'Auto Balance', 'effects', PICTURE),
				]),
				stack(doc, [missing(doc, 'picture-compression', 'Compression', 'visioPicture', PICTURE)]),
			]),
			shapeStylesGroup(doc, { label: 'Picture Styles', parts: ['line', 'effects'] }),
			arrangeGroup(doc),
			sizeGroup(doc, [
				command(doc, {
					id: 'picture-crop',
					label: 'Crop',
					icon: 'visioPicture',
					unsupported: CROP,
				}),
			]),
		]),
	);
}

/**
 * Container Format: Size (Margins and Automatic Resize listed disabled, Fit to Contents wired),
 * Container Styles (Fill, Line, Effects), Membership and Arrange.
 */
export function buildContainerFormatPanel(doc: Document, panel: HTMLElement): void {
	panel.dataset.contextual = '';
	panel.append(
		commandRow(doc, 'Container Format commands', [
			group(doc, 'Size', [
				stack(doc, [
					missing(doc, 'container-margins', 'Margins', 'grid', MARGINS),
					command(doc, {
						id: 'container-fit',
						label: 'Fit to Contents',
						icon: 'fit',
						size: 'small',
						action: { type: 'container', command: 'fit' },
					}),
					missing(doc, 'container-resize', 'Automatic Resize', 'grid', RESIZE),
				]),
				sizePosition(doc),
			]),
			shapeStylesGroup(doc, {
				label: 'Container Styles',
				parts: ['fill', 'line', 'effects'],
			}),
			group(doc, 'Membership', [
				missing(doc, 'container-lock', 'Lock Container', 'lock', LOCK),
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
