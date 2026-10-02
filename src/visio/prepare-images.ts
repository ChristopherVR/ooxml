import { inspectEmbeddedVisioMetafile } from './prepare-metafiles.js';
import { readVisioImage, type VisioImageOptions } from './media.js';
import type { VisioPackage } from './package.js';
import type { RawShape, Report } from './sheet.js';

/** Resolve assets before master expansion, so all instances share one validated byte array. */
export async function prepareImages(
	shapes: RawShape[],
	pkg: VisioPackage,
	sourcePart: string,
	report: Report,
	options?: VisioImageOptions,
): Promise<void> {
	for (const shape of shapes) {
		if (shape.deleted) continue;
		const localReport: Report = (code, message, extra) =>
			report(code, message, { part: sourcePart, shapeId: shape.id, ...extra });
		if (shape.foreignData) {
			if (shape.foreignData.getAttribute('ForeignType') === 'EnhMetaFile') {
				await inspectEmbeddedVisioMetafile(pkg, sourcePart, shape.foreignData, localReport);
			} else {
				const image = await readVisioImage(
					pkg,
					sourcePart,
					shape.foreignData,
					localReport,
					options,
				);
				if (image) shape.image = image;
			}
		}
		await prepareImages(shape.children, pkg, sourcePart, report, options);
	}
}
