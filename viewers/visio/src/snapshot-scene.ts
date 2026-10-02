import type {
	VisioDocument,
	VisioImage,
	VisioMatrix,
	VisioParagraph,
	VisioShape,
	VisioStyle,
	VisioText,
} from 'ooxml-core/visio';

/** Copy only named scalar fields, never arbitrary host properties or their getters. */
function fields<T, K extends keyof T>(source: T, keys: readonly K[]): Pick<T, K> {
	const result: Partial<T> = {};
	for (const key of keys) {
		const value = source[key];
		if (value !== null && (typeof value === 'object' || typeof value === 'function'))
			throw invalid();
		if (value !== undefined) result[key] = value;
	}
	return result as Pick<T, K>;
}
const matrix = (value: VisioMatrix): VisioMatrix => [
	value[0],
	value[1],
	value[2],
	value[3],
	value[4],
	value[5],
];
const invalid = () => new Error('The scene exceeds safe snapshot copy limits.');

/**
 * @internal Copy a canonically validated scene without retaining any renderer-used mutable input.
 * The caller validates the result before use. Independent allocation caps also bound a host that
 * changes its arrays while fields are read. Shared raster buffers are copied once per invocation.
 * Inspector metadata, connectors, layers and unknown host properties are not used by static SVG.
 */
export function copySnapshotScene(model: VisioDocument): VisioDocument {
	const counts = new Map<string, number>();
	const take = (key: string, amount: number, maximum: number): void => {
		const total = (counts.get(key) ?? 0) + amount;
		if (!Number.isSafeInteger(amount) || amount < 0 || total > maximum) throw invalid();
		counts.set(key, total);
	};
	const list = <T, U>(
		source: readonly T[],
		key: string,
		maximum: number,
		copy: (value: T) => U,
	): U[] => {
		if (!Array.isArray(source)) throw invalid();
		const length = source.length;
		take(key, length, maximum);
		const result: U[] = [];
		// Do not invoke a host array's iterator/map implementation or reread its changing length.
		for (let index = 0; index < length; index++) result.push(copy(source[index]!));
		return result;
	};
	const resources = new Map<Uint8Array, Uint8Array>();
	const image = (source: VisioImage): VisioImage => {
		const original = source.bytes;
		if (!(original instanceof Uint8Array) || original.byteLength > 8 * 1024 * 1024) throw invalid();
		let bytes = resources.get(original);
		if (!bytes) {
			take('raster bytes', original.byteLength, 64 * 1024 * 1024);
			bytes = new Uint8Array(original.byteLength);
			bytes.set(original);
			resources.set(original, bytes);
		}
		return {
			...fields(source, [
				'mimeType',
				'pixelWidth',
				'pixelHeight',
				'opacity',
				'x',
				'y',
				'width',
				'height',
			]),
			bytes,
		};
	};
	const style = (source: VisioStyle): VisioStyle => {
		const result = fields(source, [
			'fill',
			'lineColor',
			'lineWidth',
			'lineCap',
			'linePattern',
			'fillOpacity',
			'lineOpacity',
			'startArrow',
			'endArrow',
			'startArrowSize',
			'endArrowSize',
		]);
		const dash = source.lineDash;
		const gradient = source.fillGradient;
		return {
			...result,
			...(dash === undefined
				? {}
				: { lineDash: list(dash, 'dash values', 150_000, (value) => value) }),
			...(gradient
				? {
						fillGradient: {
							type: gradient.type,
							start: [gradient.start[0], gradient.start[1]] as const,
							end: [gradient.end[0], gradient.end[1]] as const,
							stops: list(gradient.stops, 'gradient stops', 100_000, (stop) =>
								fields(stop, ['offset', 'color', 'opacity']),
							),
						},
					}
				: {}),
		};
	};
	const paragraph = (source: VisioParagraph): VisioParagraph => ({
		...fields(source, [
			'start',
			'end',
			'horizontalAlign',
			'indentLeft',
			'indentRight',
			'indentFirst',
			'spaceBefore',
			'spaceAfter',
			'direction',
		]),
		lineSpacing: fields(source.lineSpacing, ['kind', 'value']),
		...(source.bullet
			? { bullet: fields(source.bullet, ['text', 'fontFamily', 'fontSize', 'offset']) }
			: {}),
	});
	const text = (source: VisioText): VisioText => ({
		...fields(source, [
			'plainText',
			'backgroundColor',
			'backgroundOpacity',
			'fontFamily',
			'fontSize',
			'color',
			'horizontalAlign',
			'verticalAlign',
			'width',
			'height',
		]),
		transform: matrix(source.transform),
		margins: fields(source.margins, ['left', 'right', 'top', 'bottom']),
		runs: list(source.runs, 'text runs', 100_000, (run) => ({
			...fields(run, ['text', 'fontFamily', 'fontSize', 'color']),
			bold: !!run.bold,
			italic: !!run.italic,
			underline: !!run.underline,
		})),
		...(source.paragraphs === undefined
			? {}
			: { paragraphs: list(source.paragraphs, 'paragraphs', 50_000, paragraph) }),
	});
	const seen = new Set<VisioShape>();
	const shapes = (source: VisioShape[], depth = 0): VisioShape[] => {
		if (depth > 64) {
			if (!Array.isArray(source) || source.length) throw invalid();
			return [];
		}
		return list(source, 'shapes', 25_000, (shape) => {
			if (seen.has(shape)) throw invalid();
			seen.add(shape);
			return {
				...fields(shape, ['id', 'name', 'kind', 'groupDisplayMode', 'width', 'height', 'hidden']),
				transform: matrix(shape.transform),
				geometry: list(shape.geometry, 'geometry', 100_000, (geometry) =>
					fields(geometry, ['path', 'fill', 'stroke']),
				),
				style: style(shape.style),
				text: text(shape.text),
				children: shapes(shape.children, depth + 1),
				...(shape.image ? { image: image(shape.image) } : {}),
			};
		});
	};
	return {
		format: 'vsdx',
		diagnostics: list(model.diagnostics, 'diagnostics', 2000, (note) =>
			fields(note, ['code', 'message', 'severity']),
		),
		pages: list(model.pages, 'pages', 256, (page) => ({
			...fields(page, ['id', 'name', 'width', 'height', 'isBackground', 'backgroundPageId']),
			shapes: shapes(page.shapes),
			connectors: [],
		})),
	};
}
