// Inserting and editing charts and pictures.
import type { ChartObject, DrawingAnchor, ImageObject, Workbook } from '../model.js';
import { type EditContext, sheetAt } from './context.js';

/** File extensions for the picture types Excel accepts. */
export const IMAGE_EXTENSIONS: Readonly<Record<string, string>> = {
	'image/png': 'png',
	'image/jpeg': 'jpeg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
	'image/tiff': 'tiff',
	'image/webp': 'webp',
	'image/svg+xml': 'svg',
	'image/x-emf': 'emf',
	'image/x-wmf': 'wmf',
};

/** What `updateChart` may change (the chart keeps its kind and source part). */
export type ChartPatch = Partial<Omit<ChartObject, 'kind' | 'partName'>>;

/** Adds a chart drawn from the model (a new chart part is written on save); returns its index. */
export function addChart(ctx: EditContext, s: number, chart: Omit<ChartObject, 'kind'>): number {
	const sheet = sheetAt(ctx.workbook, s);
	return ctx.run(
		'Insert chart',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const added: ChartObject = { ...structuredClone(chart), kind: 'chart' };
			delete added.partName;
			sheet.drawings.push(added);
			return sheet.drawings.length - 1;
		},
		{ sheet: s },
	);
}

/**
 * Edits a chart. A chart read from a file keeps its part: on save the title, legend, series
 * names and references are patched into it so formatting the model does not carry survives;
 * changing the type, grouping or number of series rewrites the part from the model.
 */
export function updateChart(ctx: EditContext, s: number, index: number, patch: ChartPatch): void {
	const sheet = sheetAt(ctx.workbook, s);
	const chart = sheet.drawings[index];
	if (!chart || chart.kind !== 'chart') throw new RangeError(`No chart at index ${index}`);
	ctx.run(
		'Edit chart',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const live = sheet.drawings[index] as ChartObject;
			for (const [key, value] of Object.entries(structuredClone(patch))) {
				if (key === 'kind' || key === 'partName') continue;
				if (value === undefined) delete (live as unknown as Record<string, unknown>)[key];
				else (live as unknown as Record<string, unknown>)[key] = value;
			}
		},
		{ sheet: s },
	);
}

/** A media part name not used by the workbook's package yet (`xl/media/imageN.ext`). */
export function newMediaPart(workbook: Workbook, contentType: string): string {
	const ext = IMAGE_EXTENSIONS[contentType] ?? contentType.split('/').pop() ?? 'bin';
	const taken = new Set([...(workbook.source?.parts.keys() ?? [])].map((p) => p.toLowerCase()));
	for (const sheet of workbook.sheets)
		for (const d of sheet.drawings) if (d.kind === 'image') taken.add(d.partName.toLowerCase());
	for (let n = 1; ; n++) {
		const name = `xl/media/image${n}.${ext}`;
		if (!taken.has(name)) return name;
	}
}

/**
 * Inserts a picture. Its bytes are stored in the workbook's package (created when the workbook
 * has none) under a new media part, which the writer copies on save. Returns the image.
 */
export function addImage(
	ctx: EditContext,
	s: number,
	bytes: Uint8Array,
	contentType: string,
	anchor: DrawingAnchor,
	name?: string,
): ImageObject {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	if (!contentType.startsWith('image/')) throw new Error(`Not an image type: ${contentType}`);
	const partName = newMediaPart(workbook, contentType);
	return ctx.run(
		'Insert picture',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			workbook.source ??= { parts: new Map() };
			workbook.source.parts.set(partName, bytes.slice());
			const image: ImageObject = {
				kind: 'image',
				anchor: structuredClone(anchor),
				partName,
				contentType,
			};
			if (name) image.name = name;
			sheet.drawings.push(image);
			return structuredClone(image);
		},
		{ sheet: s },
	);
}
