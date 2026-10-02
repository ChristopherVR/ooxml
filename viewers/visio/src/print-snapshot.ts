import { getVisioPageLayers, type VisioDocument, type VisioShape } from 'ooxml-core/visio';
import type { CompatibilityNote } from './diagnostics.js';
import { estimatePageSvgBytes, exportPageSvg, MAX_SVG_EXPORT_BYTES } from './export-svg.js';
import { assertViewableDocument } from './scene-validation.js';
import { copySnapshotScene } from './snapshot-scene.js';

/** Application guardrails, not printer capabilities or Microsoft Visio limits. */
export const PRINT_SNAPSHOT_LIMITS = Object.freeze({
	maxPages: 32,
	maxTotalBytes: 32 * 1024 * 1024,
	maxPageBytes: MAX_SVG_EXPORT_BYTES,
	maxShapeInstances: 100_000,
	maxGeometryInstances: 100_000,
	maxPathCharacters: 16_000_000,
	maxTextCharacters: 1_000_000,
	maxTextRuns: 100_000,
	maxImageBytes: 64 * 1024 * 1024,
	maxUniqueImagePixels: 64_000_000,
	maxImageInstancePixels: 256_000_000,
	maxPageAxisCssPixels: 4096,
	maxPageAreaCssPixels: 16_000_000,
	maxTotalAreaCssPixels: 64_000_000,
	maxValidationShapeVisits: 1_000_000,
	maxValidationWork: 256 * 1024 * 1024,
});
export type PrintSnapshotLimits = { -readonly [K in keyof typeof PRINT_SNAPSHOT_LIMITS]: number };
export interface PrintSnapshotOptions {
	/** Explicit output order. Background pages are included as sheets only when selected. */
	readonly pageIndices: readonly number[];
	/** Positive integer limits may lower, but never raise, the hard ceilings. */
	readonly limits?: Partial<PrintSnapshotLimits>;
}
/** Viewer handles always capture exactly the current page. */
export type CurrentPagePrintSnapshotOptions = Omit<PrintSnapshotOptions, 'pageIndices'>;
export interface PrintSnapshotPage {
	readonly pageIndex: number;
	readonly pageId: string;
	readonly pageName: string;
	/** Saved drawing dimensions, not printer paper, margins, orientation or scale. */
	readonly drawingSize: Readonly<{ width: number; height: number; unit: 'in' }>;
	/** An isolated SVG document. Do not concatenate page trees with colliding resource IDs. */
	readonly svg: string;
	readonly byteLength: number;
	readonly diagnostics: readonly Readonly<CompatibilityNote>[];
}
export interface PrintSnapshotUsage {
	readonly estimatedSvgBytes: number;
	readonly svgBytes: number;
	readonly shapeInstances: number;
	readonly geometryInstances: number;
	readonly pathCharacters: number;
	readonly textCharacters: number;
	readonly textRuns: number;
	/** Raw resources counted once per byte identity per output page, including backgrounds. */
	readonly imageBytes: number;
	readonly uniqueImagePixels: number;
	readonly imageInstancePixels: number;
	readonly totalAreaCssPixels: number;
	readonly validationShapeVisits: number;
	/** Conservative bytes, characters and records across validation and the private copy pass. */
	readonly validationWork: number;
}
export interface PrintSnapshot {
	readonly appearance: 'saved-display';
	readonly pageIndices: readonly number[];
	readonly pages: readonly PrintSnapshotPage[];
	readonly byteLength: number;
	readonly limits: Readonly<PrintSnapshotLimits>;
	readonly usage: PrintSnapshotUsage;
	readonly limitations: readonly string[];
}
const LIMITATIONS = Object.freeze([
	'This snapshot reproduces supported saved drawing display appearance, independent of viewport zoom, pan, selection, search highlighting and temporary viewer filters. It is not a guarantee of Microsoft Visio fidelity.',
	'Saved print visibility, including per-layer Print and NonPrinting, is not applied. This snapshot preserves saved display visibility and does not reveal hidden content for printing.',
	'Saved printer settings are unsupported. Drawing dimensions are not printer paper, margins, orientation, tiling or print scale. This snapshot only prepares isolated drawing artifacts and does not print.',
]);
const exceeded = (name: string): Error =>
	new Error(`Print snapshot exceeds the safe ${name} limit. Select fewer or smaller pages.`);
function within(value: number, limit: number, name: string): void {
	if (!Number.isFinite(value) || value > limit) throw exceeded(name);
}
function copyLimits(input: PrintSnapshotOptions['limits']): Readonly<PrintSnapshotLimits> {
	if (input !== undefined && (!input || typeof input !== 'object' || Array.isArray(input)))
		throw new Error('Print snapshot limits must be an object.');
	const limits: PrintSnapshotLimits = { ...PRINT_SNAPSHOT_LIMITS };
	for (const key of Object.keys(input ?? {})) {
		if (!Object.hasOwn(limits, key)) throw new Error(`Unknown print snapshot limit: ${key}.`);
		const name = key as keyof PrintSnapshotLimits;
		const value = input![name];
		if (!Number.isSafeInteger(value) || value! < 1 || value! > PRINT_SNAPSHOT_LIMITS[name])
			throw new Error(
				`Print snapshot ${name} must be a positive integer no greater than ${PRINT_SNAPSHOT_LIMITS[name]}.`,
			);
		limits[name] = value!;
	}
	return Object.freeze(limits);
}

/**
 * Synchronously prepare bounded, immutable page artifacts through the shared SVG renderer.
 * Requires its DOM/serializer, but never mounts nodes, downloads, fetches, opens frames or prints.
 * Canonical scene/raster validation is repeated on every call. No source object is frozen.
 * Conservative budgets include hidden content and can reject otherwise exportable individual pages.
 */
export function createPrintSnapshot(
	model: VisioDocument,
	options: PrintSnapshotOptions,
): PrintSnapshot {
	if (!options || typeof options !== 'object' || !Array.isArray(options.pageIndices))
		throw new Error('Print snapshot requires an explicit pageIndices array.');
	for (const key of Object.keys(options))
		if (key !== 'pageIndices' && key !== 'limits')
			throw new Error(`Unknown print snapshot option: ${key}.`);
	const limits = copyLimits(options.limits);
	if (!options.pageIndices.length)
		throw new Error('Print snapshot page selection must not be empty.');
	within(options.pageIndices.length, limits.maxPages, 'selected page count');
	const pageIndices = Object.freeze([...options.pageIndices]);
	const seen = new Set<number>();
	for (const index of pageIndices) {
		if (
			!Number.isSafeInteger(index) ||
			index < 0 ||
			!Array.isArray(model.pages) ||
			index >= model.pages.length
		)
			throw new Error('Print snapshot page index is out of range.');
		if (seen.has(index))
			throw new Error('Print snapshot page selection contains duplicate indices.');
		seen.add(index);
	}
	assertViewableDocument(model);
	const validation = validationCost(model);
	// Account conservatively for source validation, a bounded private copy, validation of that
	// copy, and exportPageSvg + renderPage validation per page, including unselected images.
	const passes = 3 + 2 * pageIndices.length;
	const usage = {
		estimatedSvgBytes: 0,
		svgBytes: 0,
		shapeInstances: 0,
		geometryInstances: 0,
		pathCharacters: 0,
		textCharacters: 0,
		textRuns: 0,
		imageBytes: 0,
		uniqueImagePixels: 0,
		imageInstancePixels: 0,
		totalAreaCssPixels: 0,
		validationShapeVisits: validation.shapes * passes,
		validationWork: validation.work * passes,
	};
	within(
		usage.validationShapeVisits,
		limits.maxValidationShapeVisits,
		'whole-document validation shape visits',
	);
	within(usage.validationWork, limits.maxValidationWork, 'whole-document validation work');
	// All later bounds, rendering and metadata read the same detached state. DOM/serializer
	// hooks can reenter a host and mutate its original document without changing this job.
	model = copySnapshotScene(model);
	assertViewableDocument(model);
	// Known-field getters may have changed the source while it was copied. Charge the larger
	// validated scene, rather than letting a cheap original estimate undercount later exports.
	const detached = validationCost(model);
	usage.validationShapeVisits = Math.max(validation.shapes, detached.shapes) * passes;
	usage.validationWork = Math.max(validation.work, detached.work) * passes;
	within(
		usage.validationShapeVisits,
		limits.maxValidationShapeVisits,
		'whole-document validation shape visits',
	);
	within(usage.validationWork, limits.maxValidationWork, 'whole-document validation work');
	const add = (key: keyof typeof usage, value: number, limit: number): void => {
		usage[key] += value;
		within(usage[key], limit, key);
	};
	// Finish the whole job's dimension, composition and allocation preflight before any export.
	for (const index of pageIndices) {
		const page = model.pages[index]!;
		const width = page.width * 96,
			height = page.height * 96;
		within(width, limits.maxPageAxisCssPixels, 'drawing width in CSS pixels');
		within(height, limits.maxPageAxisCssPixels, 'drawing height in CSS pixels');
		within(width * height, limits.maxPageAreaCssPixels, 'drawing area in CSS pixels');
		add('totalAreaCssPixels', width * height, limits.maxTotalAreaCssPixels);
		const resources = new Set<Uint8Array>();
		for (const layer of getVisioPageLayers(model, page.id)) {
			const stack = [...layer.shapes];
			while (stack.length) {
				const shape = stack.pop()!;
				add('shapeInstances', 1, limits.maxShapeInstances);
				add('geometryInstances', shape.geometry.length, limits.maxGeometryInstances);
				for (const geometry of shape.geometry)
					add('pathCharacters', geometry.path.length, limits.maxPathCharacters);
				add('textCharacters', shape.text.plainText.length, limits.maxTextCharacters);
				add('textRuns', shape.text.runs.length, limits.maxTextRuns);
				for (const run of shape.text.runs)
					add('textCharacters', run.text.length, limits.maxTextCharacters);
				if (shape.image) {
					const { bytes, pixelWidth, pixelHeight } = shape.image;
					add('imageInstancePixels', pixelWidth * pixelHeight, limits.maxImageInstancePixels);
					if (!resources.has(bytes)) {
						resources.add(bytes);
						add('imageBytes', bytes.byteLength, limits.maxImageBytes);
						add('uniqueImagePixels', pixelWidth * pixelHeight, limits.maxUniqueImagePixels);
					}
				}
				stack.push(...shape.children);
			}
		}
		add(
			'estimatedSvgBytes',
			estimatePageSvgBytes(model, page, limits.maxPageBytes),
			limits.maxTotalBytes,
		);
	}
	const pages: PrintSnapshotPage[] = [];
	for (const pageIndex of pageIndices) {
		const pageId = model.pages[pageIndex]!.id;
		const result = exportPageSvg(model, pageIndex, { maxBytes: limits.maxPageBytes });
		add('svgBytes', result.byteLength, limits.maxTotalBytes);
		pages.push(
			Object.freeze({
				pageIndex,
				pageId,
				pageName: result.pageName,
				drawingSize: Object.freeze({
					width: result.width,
					height: result.height,
					unit: 'in' as const,
				}),
				svg: result.svg,
				byteLength: result.byteLength,
				diagnostics: Object.freeze(result.diagnostics.map((note) => Object.freeze({ ...note }))),
			}),
		);
	}
	return Object.freeze({
		appearance: 'saved-display',
		pageIndices,
		pages: Object.freeze(pages),
		byteLength: usage.svgBytes,
		limits,
		usage: Object.freeze(usage),
		limitations: LIMITATIONS,
	});
}

/** Cost accounting only. Format interpretation and safety checks remain canonical. */
function validationCost(model: VisioDocument): { shapes: number; work: number } {
	let shapes = 0,
		work = model.pages.length;
	const images = new Set<Uint8Array>();
	const strings = (...values: (string | undefined)[]): void => {
		for (const value of values) work += 1 + (value?.length ?? 0);
	};
	for (const note of model.diagnostics) strings(note.code, note.message);
	for (const page of model.pages) {
		strings(page.id, page.name, page.backgroundPageId);
		for (const layer of page.layers ?? []) strings(layer.id, layer.name);
		const stack: VisioShape[] = [...page.shapes];
		while (stack.length) {
			const shape = stack.pop()!;
			shapes++;
			work += 64;
			strings(
				shape.id,
				shape.name,
				shape.style.fill,
				shape.style.lineColor,
				shape.text.plainText,
				shape.text.fontFamily,
				shape.text.color,
				shape.text.backgroundColor,
			);
			for (const geometry of shape.geometry) strings(geometry.path);
			for (const stop of shape.style.fillGradient?.stops ?? []) {
				work += 3;
				strings(stop.color);
			}
			for (const run of shape.text.runs) {
				work += 8;
				strings(run.text, run.fontFamily, run.color);
			}
			for (const paragraph of shape.text.paragraphs ?? []) {
				work += 16;
				if (paragraph.bullet) strings(paragraph.bullet.text, paragraph.bullet.fontFamily);
			}
			for (const rows of [shape.shapeData ?? [], shape.hyperlinks ?? []])
				for (const row of rows) {
					const count = (record: object): void => {
						for (const value of Object.values(record))
							work += 1 + (typeof value === 'string' ? value.length : 0);
					};
					count(row);
					if ('target' in row) count(row.target);
				}
			if (shape.image && !images.has(shape.image.bytes)) {
				images.add(shape.image.bytes);
				work += shape.image.bytes.byteLength;
			}
			stack.push(...shape.children);
		}
	}
	return { shapes, work };
}
