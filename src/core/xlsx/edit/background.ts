import { IMAGE_EXTENSIONS, newMediaPart } from './charts';
import { type EditContext, sheetAt } from './context';

export function setSheetBackground(
	ctx: EditContext,
	s: number,
	image: { bytes: Uint8Array; contentType: string } | undefined,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	if (!image && !sheet.backgroundPicture && !sheet.preserved.has('picture')) return;
	if (image && (!Object.hasOwn(IMAGE_EXTENSIONS, image.contentType) || !image.bytes.length))
		throw new RangeError('A worksheet background requires a supported image.');
	ctx.run(
		image ? 'Set sheet background' : 'Delete sheet background',
		'annotations',
		[{ kind: 'parts', sheet: s, parts: ['backgroundPicture', 'preserved'] }],
		() => {
			sheet.preserved.delete('picture');
			if (!image) {
				delete sheet.backgroundPicture;
				return;
			}
			const partName = newMediaPart(ctx.workbook, image.contentType);
			ctx.workbook.source ??= { parts: new Map() };
			ctx.workbook.source.parts.set(partName, image.bytes.slice());
			sheet.backgroundPicture = { partName, contentType: image.contentType };
		},
		{ sheet: s },
	);
}
