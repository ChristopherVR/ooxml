import { RASTER_IMAGE_LIMITS } from 'ooxml-core/drawingml';
import { anchorToPixelBox, createGridMetrics } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { field, numberInput } from '../dialogs/fields';
import { showDialog } from '../dialogs/frame';
import { pickImage } from './insert';
import { bounded, type FormatGuard } from './drawing-format-actions';
import type { DrawingFormatAction, FormatApply, FormatSelection } from './drawing-format';

const sides = [
	['cropLeft', 'Left (%)'],
	['cropTop', 'Top (%)'],
	['cropRight', 'Right (%)'],
	['cropBottom', 'Bottom (%)'],
] as const;

export async function runPictureAdjustment(
	ctx: EditorContext,
	selected: Extract<FormatSelection, { kind: 'picture' }>,
	live: FormatGuard,
	apply: FormatApply,
	id: DrawingFormatAction,
	arg: unknown,
): Promise<void> {
	const {
		current: { ws, sheet, workbook, session },
		drawing,
		index,
	} = selected;
	const editEffects = (patch: Parameters<typeof session.editPictureEffects>[2]) => {
		if (live()) session.editPictureEffects(sheet, index, patch);
	};
	const label = ctx.commands.get(id)?.label ?? '';
	switch (id) {
		case 'picture.corrections': {
			const lum = drawing.effects?.find((effect) => effect.kind === 'luminance');
			const brightness = numberInput(ctx, (lum?.brightness ?? 0) * 100, -100, 100, 0.01);
			const contrast = numberInput(ctx, (lum?.contrast ?? 0) * 100, -100, 100, 0.01);
			await showDialog(ctx, {
				name: 'picture-corrections',
				heading: label,
				body: [field(ctx, 'Brightness (%)', brightness), field(ctx, 'Contrast (%)', contrast)],
				submit: () => {
					const b = bounded(brightness),
						c = bounded(contrast);
					if (b === undefined || c === undefined || !live()) return undefined;
					editEffects({ brightness: b / 100, contrast: c / 100 });
					return true;
				},
			});
			return;
		}
		case 'picture.color':
			if (typeof arg === 'boolean') editEffects({ grayscale: arg });
			return;
		case 'picture.crop': {
			const inputs = sides.map(([key, title]) => ({
				key,
				title,
				input: numberInput(
					ctx,
					Math.round((drawing.crop?.[key] ?? 0) * 10000) / 100,
					-100,
					99,
					0.01,
				),
			}));
			await showDialog(ctx, {
				name: 'picture-crop',
				heading: label,
				body: inputs.map(({ title, input }) => field(ctx, title, input)),
				submit: () => {
					const crop = { cropLeft: 0, cropTop: 0, cropRight: 0, cropBottom: 0 };
					for (const { key, input } of inputs) {
						const percent = bounded(input);
						if (percent === undefined) return undefined;
						crop[key] = percent / 100;
					}
					if (!live()) return undefined;
					if (crop.cropLeft + crop.cropRight >= 1 || crop.cropTop + crop.cropBottom >= 1) {
						ctx.toast(ctx.t('Crop must leave a positive width and height.'), 'warning');
						return undefined;
					}
					session.setPictureCrop(sheet, index, crop);
					return true;
				},
			});
			return;
		}
		case 'picture.change': {
			const file = await pickImage(ctx, 'image/png,image/jpeg,image/gif');
			if (!file || !live()) return;
			try {
				if (file.size > RASTER_IMAGE_LIMITS.maxImageBytes) throw new RangeError();
				const bytes = new Uint8Array(await file.arrayBuffer());
				apply({ kind: 'replacePicture', bytes });
			} catch {
				ctx.toast(
					ctx.t('Choose a valid PNG, JPEG or GIF within the picture size limits.'),
					'warning',
				);
			}
			return;
		}
		case 'picture.compress': {
			const part = drawing.partName;
			const bytes = workbook.source?.parts.get(part);
			if (!bytes) return;
			const resolution = numberInput(ctx, 96, 48, 300);
			const ppi = await showDialog(ctx, {
				name: 'picture-compress',
				heading: label,
				body: field(ctx, 'Resolution (ppi)', resolution),
				submit: () => bounded(resolution),
			});
			if (ppi === undefined || !live() || drawing.partName !== part) return;
			try {
				const { compressPicture } = await import('../picture-compress');
				if (!live()) return;
				const box = anchorToPixelBox(ws, drawing.anchor, createGridMetrics(ws, { zoom: 100 }));
				const crop = drawing.crop;
				const image = await compressPicture(bytes, {
					width: box.w,
					height: box.h,
					ppi,
					deleteCrop: false,
					crop: [
						crop?.cropLeft ?? 0,
						crop?.cropTop ?? 0,
						crop?.cropRight ?? 0,
						crop?.cropBottom ?? 0,
					],
				});
				if (!live() || drawing.partName !== part) return;
				if (image) {
					const compressed = new Uint8Array(await image.arrayBuffer());
					if (drawing.partName === part) apply({ kind: 'compressPicture', bytes: compressed });
				} else {
					ctx.toast(ctx.t('The picture is already smaller than the compressed result.'), 'info');
				}
			} catch {
				ctx.toast(
					ctx.t('This picture could not be compressed within the picture size limits.'),
					'warning',
				);
			}
		}
	}
}
