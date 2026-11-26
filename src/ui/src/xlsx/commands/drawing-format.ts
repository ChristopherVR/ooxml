import { editing, target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { selectedDrawings } from '../drawing-selection';

function formatSelection(current: NonNullable<ReturnType<typeof target>>, index: number) {
	const drawing = current.ws.drawings[index];
	if (drawing?.kind === 'image' && !drawing.svgPartName)
		return { kind: 'picture' as const, current, index, drawing };
	if (drawing?.kind === 'unsupported' && drawing.shape)
		return { kind: 'shape' as const, current, index, drawing, shape: drawing.shape };
}

export function formatTabContext(ctx: EditorContext, kind: 'picture' | 'shape'): boolean {
	const indices = selectedDrawings(ctx);
	const index = indices.at(-1);
	const drawing = index === undefined ? undefined : target(ctx)?.ws.drawings[index];
	return kind === 'picture'
		? drawing?.kind === 'image' && !drawing.svgPartName
		: drawing?.kind === 'unsupported' && !!drawing.shape;
}

export type FormatSelection = NonNullable<ReturnType<typeof formatSelection>>;
export type FormatApply = (
	intent: Parameters<FormatSelection['current']['session']['formatDrawing']>[2],
) => boolean;

const multiFormatActions = new Set<DrawingFormatAction>([
	'drawing.size',
	'drawing.shape-style',
	'drawing.shape-fill',
	'drawing.shape-outline',
	'drawing.shape-effects',
	'picture.style',
	'picture.border',
	'picture.effects',
	'picture.reset',
]);

function formatTargets(ctx: EditorContext, id: DrawingFormatAction): FormatSelection[] {
	const current = target(ctx);
	const indices = selectedDrawings(ctx);
	if (!current || (indices.length !== 1 && !multiFormatActions.has(id))) return [];
	const targets: FormatSelection[] = [];
	for (const index of indices) {
		const selected = formatSelection(current, index);
		if (!selected || !canFormat(id, selected)) return [];
		targets.push(selected);
	}
	return targets;
}

const single = 'Select one picture or shape on an editable sheet.';
const multiple = 'Select pictures or shapes on an editable sheet.';
const shape = 'Select a shape that supports this formatting.';
const wordArt = 'Select a supported WordArt text box.';
const raster = 'Select a supported raster picture without an artistic cache.';
const adjustments = 'Select a picture with unambiguous adjustment settings.';
const shadow = 'Select a picture or shape with editable frame effects.';

export const drawingFormatActions = {
	'drawing.alt-text': ['Alt Text', single],
	'drawing.size': ['Size', multiple],
	'drawing.edit-shape': ['Edit Shape', shape],
	'drawing.shape-style': ['Shape Styles', shape],
	'drawing.shape-fill': ['Shape Fill', shape],
	'drawing.shape-outline': ['Shape Outline', shape],
	'drawing.shape-effects': ['Shape Effects', shadow],
	'drawing.word-art-style': ['WordArt Styles', wordArt],
	'drawing.text-fill': ['Text Fill', wordArt],
	'drawing.text-outline': ['Text Outline', wordArt],
	'drawing.text-effects': ['Text Effects', wordArt],
	'picture.corrections': ['Corrections', adjustments],
	'picture.color': ['Color', adjustments],
	'picture.artistic-effects': ['Artistic Effects', 'Artistic effects require a baked image cache.'],
	'picture.compress': [
		'Compress Pictures',
		'Select a PNG or JPEG picture and use a browser with workers and OffscreenCanvas.',
	],
	'picture.change': ['Change Picture', raster],
	'picture.reset': ['Reset Picture', shadow],
	'picture.style': ['Picture Styles', shadow],
	'picture.border': ['Picture Border', multiple],
	'picture.effects': ['Picture Effects', shadow],
	'picture.crop': ['Crop', 'Select a picture with a standard stretch layout.'],
} as const;

export type DrawingFormatAction = keyof typeof drawingFormatActions;

function canFormat(id: DrawingFormatAction, selected: FormatSelection): boolean {
	if (id === 'drawing.size') return true;
	if (selected.kind === 'shape') {
		if (id === 'picture.border') id = 'drawing.shape-outline';
		if (id === 'picture.effects') id = 'drawing.shape-effects';
		const shape = selected.shape;
		if (id.startsWith('picture.') || shape.equationParagraphs) return false;
		if (id === 'drawing.alt-text') return true;
		if (shape.kind === 'group') return false;
		switch (id) {
			case 'drawing.edit-shape':
				return shape.kind === 'shape' && !shape.textBox;
			case 'drawing.shape-effects':
				return !shape.effectDag;
			case 'drawing.word-art-style':
			case 'drawing.text-fill':
			case 'drawing.text-outline':
			case 'drawing.text-effects':
				return !!shape.wordArt && !!shape.text;
			default:
				return true;
		}
	}
	const image = selected.drawing;
	const effects = image.effects ?? [];
	const hasArtistic = effects.some((effect) => effect.kind === 'artistic');
	switch (id) {
		case 'picture.artistic-effects':
			return false;
		case 'picture.change':
			return !hasArtistic;
		case 'picture.compress':
			return (
				!hasArtistic &&
				['image/png', 'image/jpeg'].includes(image.contentType) &&
				typeof Worker !== 'undefined' &&
				typeof OffscreenCanvas !== 'undefined' &&
				typeof createImageBitmap === 'function'
			);
		case 'picture.reset':
			return true;
		case 'picture.style':
		case 'picture.effects':
		case 'drawing.shape-effects':
			return !image.effectDag;
		case 'picture.corrections':
			return (
				effects.filter(
					(effect) =>
						effect.kind === 'luminance' || (effect.kind === 'unsupported' && effect.name === 'lum'),
				).length <= 1 &&
				!effects.some((effect) => effect.kind === 'unsupported' && effect.name === 'lum')
			);
		case 'picture.color':
			return effects.filter((effect) => effect.kind === 'grayscale').length <= 1;
		case 'picture.crop':
			return (
				!image.layout?.tile && !Object.values(image.layout?.fillRect ?? {}).some((n) => n !== 0)
			);
		case 'picture.border':
		case 'drawing.shape-outline':
		case 'drawing.alt-text':
			return true;
		default:
			return false;
	}
}

export function drawingFormatCommands() {
	return Object.entries(drawingFormatActions).map(([key, [label, unavailable]]) => {
		const id = key as DrawingFormatAction;
		return editing({
			id,
			label,
			unavailable,
			lock: 'objects',
			enabled: (ctx) => formatTargets(ctx, id).length > 0,
			run: async (ctx, arg) => {
				const targets = formatTargets(ctx, id);
				const selected = targets.at(-1);
				if (!selected) return;
				const {
					current: { session, sheet, ws },
					index,
				} = selected;
				const { runDrawingFormat } = await import('./drawing-format-actions');
				const live = () => {
					const next = formatTargets(ctx, id);
					return (
						next.length === targets.length &&
						next.every(
							(item, i) =>
								item.drawing === targets[i]?.drawing &&
								item.index === targets[i]?.index &&
								item.current.session === session &&
								item.current.ws === ws,
						) &&
						ctx.commands.isEnabled(id)
					);
				};
				const apply: FormatApply = (intent) => {
					if (!live()) return false;
					if (targets.length === 1) return session.formatDrawing(sheet, index, intent);
					session.batch(ctx.t(label), () => {
						for (const item of targets) {
							if (!session.formatDrawing(sheet, item.index, intent))
								throw new Error(ctx.t(unavailable));
						}
					});
					return true;
				};
				if (live()) await runDrawingFormat(ctx, selected, live, apply, id, arg);
			},
		});
	});
}
