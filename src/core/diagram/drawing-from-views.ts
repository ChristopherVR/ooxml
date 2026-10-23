// A `DiagramDrawing` built from laid-out shape views (a reflow, `edit/reflow.ts`), so a renderer
// that draws cached drawings (`office-ui-smartart`) can draw a diagram the producing application
// never laid out: one created or edited here. Views are in pixels; the drawing is in EMU. Shapes a
// view leaves unfilled take the theme accents in turn, with light text, as a colourful SmartArt
// does; the result is an approximation and never stands for Office's layout.
import type { DrawingColor } from '../drawingml/types';
import type { DiagramShapeView } from './drawing-view';
import type { DiagramDrawing, DiagramDrawingShape } from './types';

export interface DiagramDrawingFromViewsOptions {
	/** EMU per view unit (9525 for CSS pixels). */
	emuPerUnit?: number;
	/** Scheme colours unfilled shapes cycle through; `accent1` alone by default. */
	accents?: readonly string[];
}

const scheme = (value: string): DrawingColor => ({ kind: 'scheme', value, transforms: [] });
const srgb = (hex: string): DrawingColor => ({
	kind: 'srgb',
	value: hex.replace('#', '').toUpperCase(),
	transforms: [],
});

/** The drawing the views describe, with resolved or themed colours. */
export function diagramDrawingFromViews(
	views: readonly DiagramShapeView[],
	options: DiagramDrawingFromViewsOptions = {},
): DiagramDrawing {
	const emu = options.emuPerUnit ?? 9525;
	const accents = options.accents?.length ? options.accents : ['accent1'];
	let filled = 0;
	const shapes = views.map((view): DiagramDrawingShape => {
		const themed = !view.fillColor && !view.fillNone;
		const fill: DrawingColor = view.fillColor
			? srgb(view.fillColor)
			: scheme(accents[filled++ % accents.length] as string);
		const text = view.text?.trim() ? view.text : '';
		const textColor = view.fontColor ? srgb(view.fontColor) : scheme(themed ? 'lt1' : 'tx1');
		return {
			modelId: view.id,
			frame: {
				x: Math.round(view.x * emu),
				y: Math.round(view.y * emu),
				width: Math.round(view.width * emu),
				height: Math.round(view.height * emu),
			},
			...(view.rotation ? { rotation: view.rotation } : {}),
			geometry: view.shapeType ?? 'rect',
			fill: view.fillNone ? { kind: 'none' } : { kind: 'solid', color: fill },
			line: {
				widthEmu: 12700,
				fill: { kind: 'solid', color: view.strokeColor ? srgb(view.strokeColor) : scheme('lt1') },
			},
			...(text
				? {
						text: {
							anchor: 'ctr',
							text,
							paragraphs: text.split('\n').map((line) => ({
								runs: [
									{
										text: line,
										sizePt: Math.round((view.fontSize ?? 12) * 0.75 * 10) / 10,
										color: textColor,
									},
								],
							})),
						},
					}
				: {}),
			has3d: false,
		};
	});
	return { shapes, issues: [] };
}
