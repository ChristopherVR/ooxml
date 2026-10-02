import type { VisioDocument, VisioShape } from 'ooxml-core/visio';

export const MAX_INPUT_BYTES = 32 * 1024 * 1024;
const MAX_DIMENSION = 10_000;
function finite(value: number, label: string, min = -MAX_DIMENSION, max = MAX_DIMENSION): void {
	if (!Number.isFinite(value) || value < min || value > max)
		throw new Error(`The scene has an invalid ${label}.`);
}
/** Defensive display limits also protect callers supplying their own typed scenes. */
export function assertViewableDocument(model: VisioDocument): void {
	if (
		model.format !== 'vsdx' ||
		!Array.isArray(model.pages) ||
		model.pages.length > 256 ||
		model.diagnostics.length > 2000
	)
		throw new Error('The scene has an invalid page list.');
	const seen = new Set<VisioShape>();
	let pathBytes = 0,
		textBytes = 0,
		runCount = 0,
		imageBytes = 0,
		geometryCount = 0,
		uniquePixels = 0,
		drawnPixels = 0;
	const imageResources = new Set<Uint8Array>();
	let metadataBytes = 0,
		detailCharacters = 0,
		detailRows = 0;
	const label = (value: string, limit = 4096) => {
		if (typeof value !== 'string' || value.length > limit)
			throw new Error('The scene exceeds safe metadata string limits.');
		metadataBytes += value.length;
		if (metadataBytes > 1_000_000) throw new Error('The scene exceeds aggregate metadata limits.');
	};
	for (const diagnostic of model.diagnostics) {
		label(diagnostic.code, 256);
		label(diagnostic.message);
	}

	const stack: Array<{ shape: VisioShape; depth: number }> = [];
	for (const page of model.pages) {
		label(page.id, 256);
		label(page.name);
		for (const layer of page.layers ?? []) {
			label(layer.id, 256);
			label(layer.name);
		}
		finite(page.width, 'page width', 0.001);
		finite(page.height, 'page height', 0.001);
		for (const shape of page.shapes) stack.push({ shape, depth: 0 });
	}
	while (stack.length) {
		const { shape, depth } = stack.pop()!;
		if (depth > 64 || seen.has(shape) || seen.size >= 25_000)
			throw new Error('The scene exceeds safe shape depth/count limits or contains a cycle.');
		seen.add(shape);
		label(shape.id, 256);
		label(shape.name);
		label(shape.text.fontFamily, 1024);
		label(shape.style.fill, 256);
		label(shape.style.lineColor, 256);
		label(shape.text.color, 256);
		if (shape.text.backgroundColor) label(shape.text.backgroundColor, 256);
		if (shape.text.backgroundOpacity !== undefined)
			finite(shape.text.backgroundOpacity, 'text background opacity', 0, 1);
		const details = [...(shape.shapeData ?? []), ...(shape.hyperlinks ?? [])];
		if (details.length > 1024 || (detailRows += details.length) > 100_000)
			throw new Error('The scene exceeds shape metadata row limits.');
		const detail = (value: unknown) => {
			if (typeof value === 'string') {
				if (value.length > 8192 || (detailCharacters += value.length) > 5_000_000)
					throw new Error('The scene exceeds shape metadata string limits.');
			} else if (typeof value === 'number' && !Number.isFinite(value))
				throw new Error('The scene has invalid numeric metadata.');
		};
		for (const row of details) {
			for (const value of Object.values(row)) detail(value);
			if ('target' in row) for (const value of Object.values(row.target)) detail(value);
		}
		finite(shape.width, 'shape width', 0);
		finite(shape.height, 'shape height', 0);
		if (shape.transform.length !== 6 || shape.text.transform.length !== 6)
			throw new Error('The scene has an invalid transform.');
		for (const value of [...shape.transform, ...shape.text.transform]) finite(value, 'transform');
		finite(shape.text.width, 'text width', 0);
		finite(shape.text.height, 'text height', 0);
		finite(shape.text.fontSize, 'font size', 0, 100);
		finite(shape.style.lineWidth, 'line width', 0, 100);
		finite(shape.style.fillOpacity, 'fill opacity', 0, 1);
		finite(shape.style.lineOpacity, 'line opacity', 0, 1);
		if (shape.style.fillGradient) {
			const gradient = shape.style.fillGradient;
			if (
				gradient.type !== 'linear' ||
				gradient.stops.length < 2 ||
				gradient.stops.length > 128 ||
				gradient.start.length !== 2 ||
				gradient.end.length !== 2
			)
				throw new Error('The scene has an invalid fill gradient.');
			for (const value of [...gradient.start, ...gradient.end])
				finite(value, 'gradient position', -20_000, 20_000);
			let offset = -1;
			for (const stop of gradient.stops) {
				finite(stop.offset, 'gradient stop', 0, 1);
				finite(stop.opacity, 'gradient opacity', 0, 1);
				label(stop.color, 256);
				if (stop.offset < offset) throw new Error('Gradient stops must be ordered.');
				offset = stop.offset;
			}
		}
		textBytes += shape.text.plainText.length;
		for (const run of shape.text.runs) {
			label(run.fontFamily, 1024);
			label(run.color, 256);
			textBytes += run.text.length;
			++runCount;
			finite(run.fontSize, 'run font size', 0, 100);
		}
		for (const geometry of shape.geometry) {
			pathBytes += geometry.path.length;
			geometryCount++;
		}
		if (geometryCount > 100_000) throw new Error('The scene exceeds safe geometry limits.');
		if (pathBytes > 16_000_000 || textBytes > 5_000_000 || runCount > 100_000)
			throw new Error('The scene exceeds safe path/text limits.');
		for (const value of Object.values(shape.text.margins)) finite(value, 'text margin');
		if ((shape.text.paragraphs?.length ?? 0) > 5000)
			throw new Error('The scene exceeds safe paragraph limits.');
		for (const paragraph of shape.text.paragraphs ?? []) {
			if (
				!Number.isSafeInteger(paragraph.start) ||
				!Number.isSafeInteger(paragraph.end) ||
				paragraph.start < 0 ||
				paragraph.end < paragraph.start ||
				paragraph.end > shape.text.plainText.length
			)
				throw new Error('The scene has invalid paragraph offsets.');
			for (const value of [
				paragraph.indentLeft,
				paragraph.indentRight,
				paragraph.indentFirst,
				paragraph.spaceBefore,
				paragraph.spaceAfter,
			])
				finite(value, 'paragraph position');
			finite(paragraph.lineSpacing.value, 'line spacing', 0, 100);
			if (paragraph.bullet) {
				label(paragraph.bullet.text, 1024);
				label(paragraph.bullet.fontFamily, 1024);
				finite(paragraph.bullet.fontSize, 'bullet size', 0, 100);
				finite(paragraph.bullet.offset, 'bullet offset');
			}
		}
		if (shape.image) {
			const image = shape.image;
			if (
				!['image/png', 'image/jpeg', 'image/gif'].includes(image.mimeType) ||
				!(image.bytes instanceof Uint8Array)
			)
				throw new Error('Unsupported embedded image.');
			imageBytes += image.bytes.byteLength;
			if (image.bytes.byteLength > 8 * 1024 * 1024 || imageBytes > 64 * 1024 * 1024)
				throw new Error('The scene exceeds safe image limits.');
			finite(image.pixelWidth, 'image width', 1, 16_384);
			finite(image.pixelHeight, 'image height', 1, 16_384);
			drawnPixels += image.pixelWidth * image.pixelHeight;
			if (!imageResources.has(image.bytes)) {
				imageResources.add(image.bytes);
				uniquePixels += image.pixelWidth * image.pixelHeight;
			}
			if (uniquePixels > 64_000_000 || drawnPixels > 256_000_000)
				throw new Error('The scene exceeds aggregate decoded image limits.');
			if (image.opacity !== undefined) finite(image.opacity, 'image opacity', 0, 1);
			if (image.pixelWidth * image.pixelHeight > 32_000_000)
				throw new Error('The scene exceeds safe image pixel limits.');
			if (image.x !== undefined) finite(image.x, 'image position');
			if (image.y !== undefined) finite(image.y, 'image position');
			if (image.width !== undefined) finite(image.width, 'image width', 0);
			if (image.height !== undefined) finite(image.height, 'image height', 0);
		}
		for (const child of shape.children) stack.push({ shape: child, depth: depth + 1 });
	}
}
