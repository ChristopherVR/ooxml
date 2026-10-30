import { arrangeGroup, shapeStylesGroup } from './ribbon-arrange';
import { command, commandRow, group, menu, stack } from './ribbon-parts';
import { sizeFields } from './ribbon-size-fields';

const WORDART =
	'Text fill, outline and effects are not stored per run yet; use Home > Font for the text colour and style.';

/**
 * Visio's contextual Shape Format tab: Shape Styles and Arrange are the Home groups built again
 * (the viewer mirrors their state from Home, so one implementation serves both), Size opens the
 * Size & Position window, and WordArt Styles lists what Visio offers with the reason it is
 * unavailable here. The tab shows only while shapes are selected (`viewer-contextual-tabs.ts`).
 */
export function buildShapeFormatPanel(doc: Document, panel: HTMLElement): void {
	/** A WordArt Styles command shown, disabled, with the reason it is missing. */
	const wordArt = (id: string, label: string, icon: string) =>
		menu(doc, { id, label, icon, size: 'small', unsupported: WORDART, items: [] });
	panel.dataset.contextual = '';
	panel.append(
		commandRow(doc, 'Shape Format commands', [
			shapeStylesGroup(doc),
			group(doc, 'WordArt Styles', [
				stack(doc, [
					wordArt('text-fill', 'Text Fill', 'fill'),
					wordArt('text-outline', 'Text Outline', 'line'),
					wordArt('text-effects', 'Text Effects', 'effects'),
				]),
			]),
			arrangeGroup(doc),
			group(doc, 'Size', [
				sizeFields(doc),
				command(doc, {
					id: 'size-position',
					label: 'Size & Position',
					icon: 'sizePosition',
					action: { type: 'sizePosition' },
					checked: false,
				}),
			]),
		]),
	);
}
