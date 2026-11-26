import { WORD_ART_STYLES, wordArtRgb } from 'ooxml-core/drawingml';
import { anchorToPixelBox, createGridMetrics, shapeColor } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, numberInput, textArea } from '../dialogs/fields';
import { showDialog } from '../dialogs/frame';
import { shapeFormatStyles, pictureFormatStyles } from './drawing-format-styles';
import type { DrawingFormatAction, FormatSelection } from './drawing-format';

export type FormatApply = (
	intent: Parameters<FormatSelection['current']['session']['formatDrawing']>[2],
) => boolean;
export type FormatGuard = () => boolean;

export function bounded(input: HTMLInputElement): number | undefined {
	const value = input.valueAsNumber;
	return Number.isFinite(value) && input.validity.valid ? value : undefined;
}

async function paint(
	ctx: EditorContext,
	label: string,
	initial: string | undefined,
	allowNone: boolean,
) {
	const color = ctx.host.ownerDocument.createElement('input');
	color.type = 'color';
	color.value = '#' + (wordArtRgb(initial?.replace(/^#/, '')) ?? '000000');
	const none = checkbox(ctx, 'None', allowNone && !initial);
	color.disabled = none.input.checked;
	none.input.addEventListener('change', () => {
		color.disabled = none.input.checked;
	});
	return showDialog(ctx, {
		name: 'drawing-color',
		heading: label,
		body: [field(ctx, 'Color', color), ...(allowNone ? [none.wrapper] : [])],
		submit: () => (none.input.checked ? null : color.value.slice(1)),
	});
}

const presetIndex = (arg: unknown) =>
	typeof arg === 'string' && /^\d{1,2}$/.test(arg) ? Number(arg) : -1;

export async function runDrawingFormat(
	ctx: EditorContext,
	selected: FormatSelection,
	live: FormatGuard,
	id: DrawingFormatAction,
	arg: unknown,
): Promise<void> {
	const {
		current: { ws, sheet, session, workbook },
		drawing,
		index,
	} = selected;
	const apply: FormatApply = (intent) => live() && session.formatDrawing(sheet, index, intent);
	const model = selected.kind === 'shape' ? selected.shape : selected.drawing;
	const label = ctx.commands.get(id)?.label ?? '';
	switch (id) {
		case 'drawing.alt-text': {
			const input = textArea(ctx, model.description ?? '', 4);
			input.maxLength = 4096;
			await showDialog(ctx, {
				name: 'drawing-alt-text',
				heading: label,
				body: field(ctx, 'Description', input),
				submit: () => apply({ kind: 'altText', text: input.value }) || undefined,
			});
			return;
		}
		case 'drawing.size': {
			const box = anchorToPixelBox(ws, drawing.anchor, createGridMetrics(ws, { zoom: 100 }));
			const width = numberInput(ctx, Math.round(box.w), 4, 10000);
			const height = numberInput(ctx, Math.round(box.h), 4, 10000);
			await showDialog(ctx, {
				name: 'drawing-size',
				heading: label,
				body: [field(ctx, 'Width (px)', width), field(ctx, 'Height (px)', height)],
				submit: () => {
					const w = bounded(width),
						h = bounded(height);
					return (
						(w !== undefined && h !== undefined && apply({ kind: 'size', width: w, height: h })) ||
						undefined
					);
				},
			});
			return;
		}
		case 'drawing.shape-style': {
			const style = shapeFormatStyles[presetIndex(arg)];
			if (style) apply({ kind: 'shapeStyle', color: style.color, outlined: style.outlined });
			return;
		}
		case 'picture.style': {
			const style = pictureFormatStyles[presetIndex(arg)];
			if (style) apply({ kind: 'pictureStyle', border: style.border, shadow: style.shadow });
			return;
		}
		case 'drawing.edit-shape':
			if (typeof arg === 'string') apply({ kind: 'geometry', preset: arg });
			return;
		case 'drawing.shape-fill':
		case 'drawing.shape-outline':
		case 'picture.border': {
			const active =
				id === 'drawing.shape-fill' && selected.kind === 'shape'
					? selected.shape.fill
					: model.line?.fill;
			const initial =
				active?.kind === 'solid' ? shapeColor(active.color, workbook.theme) : undefined;
			const color = await paint(ctx, label, initial, true);
			if (color !== undefined)
				apply({ kind: id === 'drawing.shape-fill' ? 'fill' : 'outline', color });
			return;
		}
		case 'drawing.shape-effects':
		case 'picture.effects':
			if (typeof arg === 'boolean') apply({ kind: 'shadow', enabled: arg });
			return;
		case 'picture.reset':
			apply({ kind: 'resetPicture' });
			return;
		case 'drawing.word-art-style':
		case 'drawing.text-fill':
		case 'drawing.text-outline':
		case 'drawing.text-effects': {
			if (selected.kind !== 'shape' || !selected.shape.wordArt || !selected.shape.text) return;
			const { wordArt: previous, text } = selected.shape;
			let style = { ...previous };
			if (id === 'drawing.word-art-style') {
				const preset = WORD_ART_STYLES[presetIndex(arg)];
				if (!preset) return;
				style = { ...preset, ...(previous.warp ? { warp: previous.warp } : {}) };
			} else if (id === 'drawing.text-effects') {
				if (arg === true) style.shadow = '404040';
				else if (arg === false) delete style.shadow;
				else return;
			} else {
				const isFill = id === 'drawing.text-fill';
				const color = await paint(ctx, label, isFill ? previous.fill : previous.outline, !isFill);
				if (color === undefined) return;
				const rgb = wordArtRgb(color);
				if (isFill && rgb) {
					style.fill = rgb;
					delete style.gradient;
				} else if (rgb) style.outline = rgb;
				else delete style.outline;
			}
			if (live()) session.editWordArt(sheet, index, { text: text.text, style });
			return;
		}
		default:
			if (selected.kind === 'picture') {
				const { runPictureAdjustment } = await import('./picture-adjustment');
				if (live()) await runPictureAdjustment(ctx, selected, live, apply, id, arg);
			}
	}
}
