/** Legacy VSD import uses the shared binary codec, never an OOXML conversion. */
import { parseVsd, type VsdDocument } from '@christophervr/ole2';
import { transform } from './geometry.js';
import { diagnosticCollector } from './diagnostics.js';
import { VisioPackageError } from './package.js';
import type { ParseVsdxOptions } from './parser.js';
import type { VisioDocument, VisioGeometry, VisioShape, VisioText } from './model.js';

export type ParseLegacyVsdOptions = Pick<
	ParseVsdxOptions,
	'maxShapes' | 'maxGeometryCommands' | 'maxTextCharacters' | 'maxDiagnostics'
> & {
	limits?: Pick<NonNullable<ParseVsdxOptions['limits']>, 'maxInputBytes' | 'maxRuntimeMs'>;
};

function limit(value: number | undefined, fallback: number): number {
	const resolved = value ?? fallback;
	if (!Number.isSafeInteger(resolved) || resolved <= 0)
		throw new VisioPackageError(
			'INVALID_LIMIT',
			'Legacy Visio limits must be positive safe integers.',
		);
	return resolved;
}

/** Normalize explicit v11 values. Unresolved styles are visibly diagnosed.
 * Missing geometry, group ownership and foreign objects are not guessed.
 */
function normalizeLegacyVsd(
	document: VsdDocument,
	options: ParseLegacyVsdOptions = {},
): VisioDocument {
	const maxShapes = limit(options.maxShapes, 25_000);
	const maxGeometry = limit(options.maxGeometryCommands, 500_000);
	const maxText = limit(options.maxTextCharacters, 5_000_000);
	const diagnostics = diagnosticCollector(limit(options.maxDiagnostics, 2_000));
	const report = diagnostics.report;
	let shapeCount = 0,
		geometryCount = 0,
		textCount = 0;
	report(
		'legacy-vsd-stored-values',
		'Legacy VSD v11 preview uses explicit stored values; masters, styles, layers, connections and background assignments are unresolved.',
	);
	const pages = document.pages.map((page) => {
		const pageId = String(page.id);
		if (page.scale !== 1)
			throw new VisioPackageError(
				'LEGACY_PAGE_SCALE',
				`Legacy page ${pageId} has an unresolved or unsupported drawing scale.`,
			);
		if (
			!Number.isFinite(page.width) ||
			!Number.isFinite(page.height) ||
			!(page.width! > 0) ||
			!(page.height! > 0)
		)
			throw new VisioPackageError(
				'LEGACY_PAGE_GEOMETRY',
				`Legacy page ${pageId} has no usable explicit dimensions.`,
			);
		const shapes: VisioShape[] = page.shapes.map((shape) => {
			const shapeId = String(shape.id),
				context = { pageId, shapeId };
			if (++shapeCount > maxShapes)
				throw new VisioPackageError('LIMIT_SHAPES', 'Legacy shape limit exceeded.');
			if (shape.kind !== 'shape')
				throw new VisioPackageError(
					'LEGACY_SHAPE_KIND',
					`Legacy shape ${shapeId} requires unresolved group ownership or foreign content.`,
				);
			if ([shape.parentId, shape.masterPageId, shape.masterShapeId].some((id) => id !== 0xffffffff))
				throw new VisioPackageError(
					'LEGACY_SHAPE_DEPENDENCY',
					`Legacy shape ${shapeId} depends on an unresolved parent or master.`,
				);
			const t = shape.transform;
			if (
				!t ||
				shape.coordinateSpace !== 'shape-local' ||
				shape.transformCoordinateSpace !== 'parent-local' ||
				![t.pinX, t.pinY, t.width, t.height, t.localPinX, t.localPinY, t.angle].every(
					Number.isFinite,
				) ||
				t.width < 0 ||
				t.height < 0
			)
				throw new VisioPackageError(
					'LEGACY_SHAPE_GEOMETRY',
					`Legacy shape ${shapeId} has no usable explicit transform.`,
				);
			const rows = shape.geometry;
			geometryCount += rows.length;
			if (geometryCount > maxGeometry)
				throw new VisioPackageError('LIMIT_GEOMETRY', 'Legacy geometry limit exceeded.');
			if (
				shape.unsupportedGeometry !== false ||
				!rows.length ||
				rows[0]?.kind !== 'moveTo' ||
				rows.some(
					(row) =>
						!['moveTo', 'lineTo'].includes(row.kind) ||
						!Number.isFinite(row.x) ||
						!Number.isFinite(row.y),
				)
			)
				throw new VisioPackageError(
					'LEGACY_SHAPE_GEOMETRY',
					`Legacy shape ${shapeId} has no supported explicit move/line geometry.`,
				);
			const geometry: VisioGeometry[] = [
				{
					path: rows
						.map((row) => `${row.kind === 'moveTo' ? 'M' : 'L'} ${row.x} ${row.y}`)
						.join(' '),
					fill: false,
					stroke: true,
				},
			];
			const plainText = shape.text ?? '';
			if (/[\u0000-\u0008\u000b-\u001f\u007f\ufffc]/u.test(plainText))
				throw new VisioPackageError(
					'LEGACY_TEXT_CONTROL',
					`Legacy shape ${shapeId} contains unresolved text fields or controls.`,
				);
			textCount += plainText.length;
			if (textCount > maxText)
				throw new VisioPackageError('LIMIT_TEXT', 'Legacy text limit exceeded.');
			report(
				'legacy-vsd-style-fallback',
				'Shape fill, line style and visibility are unresolved; preview uses an unfilled black outline.',
				context,
			);
			if (plainText)
				report(
					'legacy-vsd-text-fallback',
					'Text uses plain 12pt Arial in the shape bounds; character styles and the text-block transform are unresolved.',
					context,
				);
			const text: VisioText = {
				plainText,
				runs: plainText
					? [
							{
								text: plainText,
								fontFamily: 'Arial',
								fontSize: 1 / 6,
								color: '#000000',
								bold: false,
								italic: false,
								underline: false,
							},
						]
					: [],
				fontFamily: 'Arial',
				fontSize: 1 / 6,
				color: '#000000',
				horizontalAlign: 'left',
				verticalAlign: 'middle',
				transform: [1, 0, 0, 1, 0, 0],
				width: t.width,
				height: t.height,
				margins: { left: 0, right: 0, top: 0, bottom: 0 },
			};
			const matrix = transform(t.pinX, t.pinY, t.localPinX, t.localPinY, t.angle, t.flipX, t.flipY);
			if (!matrix.every(Number.isFinite))
				throw new VisioPackageError(
					'LEGACY_SHAPE_GEOMETRY',
					`Legacy shape ${shapeId} transform overflows.`,
				);
			return {
				id: shapeId,
				name: `Shape ${shapeId}`,
				kind: 'shape',
				width: t.width,
				height: t.height,
				transform: matrix,
				geometry,
				style: {
					fill: 'none',
					lineColor: '#000000',
					lineWidth: 0.01,
					linePattern: 1,
					fillOpacity: 1,
					lineOpacity: 1,
					startArrow: 0,
					endArrow: 0,
				},
				text,
				hidden: false,
				children: [],
			};
		});
		return {
			id: pageId,
			name: `Page ${pageId}`,
			width: page.width!,
			height: page.height!,
			isBackground: page.background,
			shapes,
			connectors: [],
		};
	});
	if (!pages.length)
		throw new VisioPackageError('LEGACY_PAGE_GEOMETRY', 'Legacy drawing has no decoded pages.');
	return { format: 'vsd', pages, diagnostics: diagnostics.finish() };
}

/** Read-only preview. The shared decoder performs bounded binary traversal.
 * Runtime checks surround synchronous decoding; they cannot preempt that call.
 */
export function parseLegacyVsd(
	input: Uint8Array | ArrayBuffer,
	options: ParseLegacyVsdOptions = {},
): VisioDocument {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	const maxInput = limit(options.limits?.maxInputBytes, 32 * 1024 * 1024);
	const maxRuntime = limit(options.limits?.maxRuntimeMs, 10_000),
		start = Date.now();
	if (bytes.byteLength > maxInput)
		throw new VisioPackageError('LIMIT_INPUT', 'Legacy Visio input exceeds limit.');
	try {
		const document = parseVsd(bytes);
		const scene = normalizeLegacyVsd(document, options);
		if (Date.now() - start >= maxRuntime)
			throw new VisioPackageError('LIMIT_RUNTIME', 'Legacy Visio processing deadline exceeded.');
		return scene;
	} catch (error) {
		if (error instanceof VisioPackageError) throw error;
		throw new VisioPackageError(
			'INVALID_LEGACY_VSD',
			`Cannot decode legacy VSD: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}
