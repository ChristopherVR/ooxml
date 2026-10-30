import type { VisioMetafileTreeConverter } from './convert-metafile';
import { connections, indexedPart, related, validateBackgrounds, visioXml } from './parts';
import { prepareImages } from './prepare-images';
import type { VisioImageOptions } from './media';
import { createLayerBudget, indexLayers, pageLayers } from './layers';
import { visioPageTheme } from './page-theme';
import type { Resources } from './style';
import { elements } from '../xml/index';
import type { VisioDocument, VisioMaster, VisioPage } from './model';
import { diagnosticCollector } from './diagnostics';
import {
	type DocumentLevel,
	type VisioDocumentCache,
	cacheMatches,
	optionsKey,
	readDocumentLevel,
} from './document-level';
import { readVisioAnnotations, readVisioComments, type VisioComment } from './comments';
import { metadata } from './metadata';
import {
	alternativeTextMetadata,
	createMetadataBudget,
	type VisioMetadataOptions,
} from './shape-metadata';
import { readDataRecordsets, type VisioDataRecordset } from './data-recordsets';
import { VisioPackage, VisioPackageError, type VisioPackageLimits } from './package';
import { normalizeShapes, type ShapeContext } from './shapes';
import { normalizeVisioPageGeometry, visioPageGeometryScale } from './page-scale';
import { readVisioPageSetup } from './page-setup';
import { readVisioPageLayout } from './page-layout';
import { readVisioSnapGlue } from './snap-glue';
import { styleSheet } from './style-inheritance';
import {
	attribute,
	child,
	children,
	number,
	readShapes,
	readSheet,
	reportCachedErrors,
	yes,
	type Report,
} from './sheet';
import { listVisioMasters } from './master-list';
import { readVisioStencilWindows } from './stencil-windows';
import { readVisioFieldProperties } from './text-field-context';
import { visioWallClock } from './text-fields';

export interface ParseVsdxOptions {
	/** Trusted converter package injection. Only supply inside a disposable, externally timed worker. */
	metafileConverter?: VisioMetafileTreeConverter;
	limits?: Partial<VisioPackageLimits>;
	images?: VisioImageOptions;
	metadata?: VisioMetadataOptions;
	maxShapes?: number;
	maxShapeDepth?: number;
	/** Aggregate weighted knot-refinement and curve subdivision work. */
	maxCurveWork?: number;
	maxGeometryCommands?: number;
	maxTextCharacters?: number;
	maxTextRuns?: number;
	maxParagraphs?: number;
	maxDiagnostics?: number;
	/**
	 * Read a stencil or template package (.vssx, .vstx) for its masters: its main part may carry
	 * the stencil or template content type, and a missing pages part means no pages. Macro-enabled
	 * packages stay refused.
	 */
	stencil?: boolean;
	/**
	 * Read only the pages with these ids (every other page is skipped after its id is checked).
	 * The result lists just those pages and leaves out what is not per page (comments, data
	 * recordsets, stencil windows, background validation); {@link mergeVisioPages} splices it into
	 * the document read earlier from a package in which only those page parts changed.
	 */
	pages?: readonly string[];
	/**
	 * Reuse the styles, themes and masters of an earlier read of this drawing when none of the
	 * parts they were read from changed (see {@link VisioDocumentCache}). The result is the one a
	 * read without it gives.
	 */
	documentCache?: VisioDocumentCache;
}

const STENCIL_MAIN_TYPES = new Set([
	'application/vnd.ms-visio.stencil.main+xml',
	'application/vnd.ms-visio.template.main+xml',
]);
function option(value: number | undefined, fallback: number): number {
	if (value === undefined) return fallback;
	if (!Number.isSafeInteger(value) || value <= 0)
		throw new VisioPackageError('INVALID_LIMIT', 'Parser limits must be positive safe integers.');
	return value;
}
/** Read-only VSDX import. This uses cached ShapeSheet values, not a formula engine. */
export async function parseVsdx(
	input: Uint8Array | ArrayBuffer,
	options: ParseVsdxOptions = {},
): Promise<VisioDocument> {
	const deadline = Date.now() + (options.limits?.maxRuntimeMs ?? 10_000);
	const checkTime = () => {
		if (Date.now() >= deadline)
			throw new VisioPackageError('LIMIT_RUNTIME', 'Visio parsing deadline exceeded.');
	};
	const maxShapes = option(options.maxShapes, 25_000),
		maxShapeDepth = option(options.maxShapeDepth, 64);
	const maxCurveWork = option(options.maxCurveWork, 20_000_000);
	let curveWork = 0;
	const maxGeometry = option(options.maxGeometryCommands, 500_000),
		maxText = option(options.maxTextCharacters, 5_000_000);
	const maxDiagnostics = option(options.maxDiagnostics, 2_000);
	const maxRuns = option(options.maxTextRuns, 100_000),
		maxParagraphs = option(options.maxParagraphs, 50_000);
	const diagnosticState = diagnosticCollector(maxDiagnostics);
	const report = diagnosticState.report;
	const pkg = await VisioPackage.open(input, options.limits);
	const types = await pkg.readXml('[Content_Types].xml', 'Types');
	if (types.namespaceURI !== 'http://schemas.openxmlformats.org/package/2006/content-types')
		throw new VisioPackageError('INVALID_CONTENT_TYPES', 'Invalid content type namespace.');
	const documentPart = await related(pkg, '', 'document');
	if (!documentPart)
		throw new VisioPackageError('INVALID_DOCUMENT', 'Missing Visio document part.');
	const documentType = elements(types)
		.find(
			(node) => node.localName === 'Override' && attribute(node, 'PartName') === `/${documentPart}`,
		)
		?.getAttribute('ContentType');
	if (
		documentType !== 'application/vnd.ms-visio.drawing.main+xml' &&
		!(options.stencil && documentType && STENCIL_MAIN_TYPES.has(documentType))
	)
		throw new VisioPackageError(
			'UNSUPPORTED_FORMAT',
			'Only VSDX drawing packages are supported; VSD, VDX, VSDM and templates are not supported.',
		);
	// Validate every relationship part, including unused ones; never dereference external URLs.
	for (const path of pkg.paths()) {
		if (path !== '_rels/.rels' && !/(^|\/)_rels\/[^/]+\.rels$/.test(path)) continue;
		const source =
			path === '_rels/.rels' ? '' : path.replace(/(^|\/)\_rels\/([^/]+)\.rels$/, '$1$2');
		for (const rel of (await pkg.relationships(source)).values()) {
			if (rel.mode === 'External')
				report(
					'external-resource-ignored',
					'External relationships are retained only as metadata and are never loaded.',
					{ part: path },
				);
		}
	}
	const cache = options.documentCache;
	const cached = cacheMatches(cache?.entry, pkg, options) ? cache!.entry : undefined;
	const stopRecording = cache && !cached ? pkg.record() : undefined;
	let level: DocumentLevel;
	if (cached) {
		level = cached.level;
		for (const args of level.reports) report(...args);
	} else level = await readDocumentLevel(pkg, documentPart, options, checkTime, report);
	const { documentRoot, styles, readPrintSettings, masters, masterRecords } = level;
	const resources: Resources = { ...level.resources };
	delete resources.pageCells;
	delete resources.fieldContext;
	let geometryCount = 0,
		textCount = 0,
		expansionCount = 0,
		textRuns = 0,
		paragraphCount = 0;
	const context: ShapeContext = {
		metadataBudget: createMetadataBudget(options.metadata),
		layerBudget: createLayerBudget(),
		masters,
		layers: new Map(),
		styles,
		styleCache: new Map(),
		defaultStyles: new Map(
			['LineStyle', 'FillStyle', 'TextStyle'].flatMap((category) => {
				const id =
					attribute(child(documentRoot, 'DocumentSheet'), category) ??
					(styles.has('0') ? '0' : undefined);
				return id === undefined ? [] : [[category, id]];
			}),
		),
		resources,
		report,
		maxShapes,
		maxShapeDepth,
		shapeCount: 0,
		checkTime,
		consumeExpansion: () => {
			checkTime();
			if (++expansionCount > maxShapes)
				throw new VisioPackageError('SHAPE_LIMIT', 'Inherited shape expansion limit exceeded.');
		},
		consumeCurveWork: (units) => {
			checkTime();
			curveWork += units;
			if (curveWork > maxCurveWork)
				throw new VisioPackageError(
					'CURVE_WORK_LIMIT',
					'Curve refinement/subdivision work limit exceeded.',
				);
		},
		consumeGeometry: () => {
			checkTime();
			if (++geometryCount > maxGeometry)
				throw new VisioPackageError('GEOMETRY_LIMIT', 'Geometry command limit exceeded.');
		},
		consumeParagraph: () => {
			checkTime();
			if (++paragraphCount > maxParagraphs)
				throw new VisioPackageError('PARAGRAPH_LIMIT', 'Paragraph count limit exceeded.');
		},
		consumeText: (length) => {
			checkTime();
			if (++textRuns > maxRuns)
				throw new VisioPackageError('TEXT_RUN_LIMIT', 'Text run count limit exceeded.');
			textCount += length;
			if (textCount > maxText)
				throw new VisioPackageError('TEXT_LIMIT', 'Text character limit exceeded.');
		},
	};
	let documentMasters: VisioMaster[];
	if (cached) {
		const { counters } = cached.listed;
		curveWork = counters.curveWork;
		geometryCount = counters.geometry;
		textCount = counters.text;
		expansionCount = counters.expansion;
		textRuns = counters.textRuns;
		paragraphCount = counters.paragraphs;
		context.shapeCount = counters.shapes;
		context.metadataBudget.restore(counters.metadata);
		context.layerBudget.restore(counters.layers);
		documentMasters = cached.listed.documentMasters;
	} else {
		documentMasters = listVisioMasters(masterRecords, context, {
			metadataBudget: createMetadataBudget(options.metadata),
			layerBudget: createLayerBudget(),
		});
		if (cache && stopRecording)
			cache.entry = {
				options: optionsKey(options),
				converter: options.metafileConverter,
				reads: stopRecording(),
				level: { ...level, resources: { ...resources } },
				listed: {
					documentMasters,
					counters: {
						curveWork,
						geometry: geometryCount,
						text: textCount,
						expansion: expansionCount,
						textRuns,
						paragraphs: paragraphCount,
						shapes: context.shapeCount,
						metadata: context.metadataBudget.state(),
						layers: context.layerBudget.state(),
					},
				},
			};
	}
	const pagesPart = await related(pkg, documentPart, 'pages', !options.stencil);
	if (!pagesPart && !options.stencil)
		throw new VisioPackageError('INVALID_DOCUMENT', 'Missing Visio pages part.');
	const pages: VisioPage[] = [],
		pageIds = new Set<string>();
	const comments: VisioComment[] = [];
	const only = options.pages ? new Set(options.pages) : undefined;
	const pageNodes = pagesPart ? children(await visioXml(pkg, pagesPart, 'Pages'), 'Page') : [];
	const foreground = pageNodes.filter((page) => !yes(attribute(page, 'Background')));
	const fieldProperties = await readVisioFieldProperties(pkg);
	const now = visioWallClock(new Date());
	for (const page of pageNodes) {
		const id = metadata(attribute(page, 'ID') ?? '', 256, 'Page ID');
		if (!id || pageIds.has(id))
			throw new VisioPackageError('INVALID_PAGE_ID', 'Page IDs must be present and unique.');
		pageIds.add(id);
		if (only && !only.has(id)) continue;
		const part = await indexedPart(pkg, pagesPart!, page, 'page');
		const root = await visioXml(pkg, part, 'PageContents');
		const localReport: Report = (code, message, extra) =>
			report(code, message, { part, pageId: id, ...extra });
		context.report = localReport;
		const sheet = readSheet(child(page, 'PageSheet'));
		reportCachedErrors(sheet, localReport);
		// PageSheet theme selectors also inherit its explicitly referenced styles.
		const pageCells = new Map<string, { value?: string; formula?: string }>();
		for (const category of ['LineStyle', 'FillStyle', 'TextStyle']) {
			const styleId = attribute(child(page, 'PageSheet'), category);
			if (styleId === undefined) continue;
			for (const [name, cell] of styleSheet(styleId, category, context).cells)
				if (
					/^(ColorSchemeIndex|EffectSchemeIndex|ConnectorSchemeIndex|FontSchemeIndex|VariationColorIndex|VariationStyleIndex)$/.test(
						name,
					)
				)
					pageCells.set(name, cell);
		}
		resources.pageCells = new Map([...pageCells, ...sheet.cells]);
		const theme = visioPageTheme(resources.themes, resources.pageCells);
		comments.push(...readVisioAnnotations(sheet, id));
		const pageNumber = foreground.indexOf(page) + 1;
		resources.fieldContext = {
			pageName: attribute(page, 'Name') ?? attribute(page, 'NameU') ?? `Page ${id}`,
			...(pageNumber ? { pageNumber } : {}),
			pageCount: foreground.length,
			properties: fieldProperties,
			now,
		};
		const layers = pageLayers(sheet, resources, localReport);
		context.layers = indexLayers(layers);
		const drawingToPageScale = visioPageGeometryScale(sheet.cells, localReport);
		const pageWidth = number(sheet.cells, 'PageWidth', 8.5, localReport),
			pageHeight = number(sheet.cells, 'PageHeight', 11, localReport);
		const width = pageWidth * drawingToPageScale,
			height = pageHeight * drawingToPageScale;
		if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0)
			throw new VisioPackageError('INVALID_PAGE_SIZE', 'Page dimensions must be positive.');
		// Shape geometry is still in drawing units here; normalizeVisioPageGeometry scales it later.
		context.pageSize = { width: pageWidth, height: pageHeight };
		const backgroundPageId = attribute(page, 'BackPage');
		if (backgroundPageId !== undefined) metadata(backgroundPageId, 256, 'Background page ID');
		const rawShapes = readShapes(root);
		await prepareImages(
			rawShapes,
			pkg,
			part,
			localReport,
			options.images,
			options.metafileConverter,
		);
		const shapes = normalizeShapes(rawShapes, context);
		normalizeVisioPageGeometry(shapes, drawingToPageScale, checkTime);
		const seen = new Set<string>();
		const visit = (items: typeof shapes): void => {
			for (const shape of items) {
				if (!shape.id || seen.has(shape.id))
					throw new VisioPackageError(
						'INVALID_SHAPE_ID',
						'Shape IDs must be present and unique on a page.',
					);
				seen.add(shape.id);
				visit(shape.children);
			}
		};
		visit(shapes);
		const connectors = connections(root, localReport);
		const pageModes: Pick<VisioPage, 'drawingSizeType' | 'drawingResizeType'> = {};
		for (const [name, key] of [
			['DrawingSizeType', 'drawingSizeType'],
			['DrawingResizeType', 'drawingResizeType'],
		] as const) {
			const source = sheet.cells.get(name);
			if (
				!source ||
				source.error !== undefined ||
				source.formula === 'Inh' ||
				(source.unit !== undefined && !['NUM', 'BOOL'].includes(source.unit))
			)
				continue;
			const value = number(sheet.cells, name, Number.NaN, localReport);
			if (Number.isSafeInteger(value) && value >= 0 && value <= 255) pageModes[key] = value;
		}
		const { paperKind, printPageOrientation, margins, printZoom, ...countMetadata } =
			readPrintSettings(children(page, 'PageSheet'));
		const pageSetup = {
			...readVisioPageSetup(sheet.cells),
			...countMetadata,
			printCountSettings: {
				...(paperKind === undefined ? {} : { paperKind }),
				...(printPageOrientation === undefined ? {} : { printPageOrientation }),
				...(margins === undefined ? {} : { margins }),
				...(printZoom === undefined ? {} : { printZoom }),
			},
		};
		const layout = readVisioPageLayout(sheet.cells);
		for (const connection of connectors)
			if (!seen.has(connection.fromShapeId) || !seen.has(connection.toShapeId))
				localReport('dangling-connection', 'A connection references a missing shape.');
		pages.push({
			id,
			...alternativeTextMetadata(sheet, localReport, context.metadataBudget),
			...pageModes,
			...(drawingToPageScale === 1 ? {} : { drawingToPageScale }),
			...(pageSetup ? { pageSetup } : {}),
			...(layout ? { layout } : {}),
			name: metadata(
				attribute(page, 'Name') ?? attribute(page, 'NameU') ?? `Page ${id}`,
				4096,
				'Page name',
			),
			width,
			height,
			isBackground: yes(attribute(page, 'Background')),
			...(backgroundPageId && backgroundPageId !== '-1' ? { backgroundPageId } : {}),
			shapes,
			connectors,
			layers,
			...(theme ? { theme } : {}),
		});
	}
	checkTime();
	if (!only) validateBackgrounds(pages, report);
	try {
		if (!only) comments.push(...(await readVisioComments(pkg, documentPart, report)));
	} catch (error) {
		if (!(error instanceof VisioPackageError) || error.code === 'LIMIT_RUNTIME') throw error;
		report('invalid-comments', 'The comments part could not be read; comments are not shown.');
	}
	let dataRecordsets: VisioDataRecordset[] = [];
	try {
		if (!only) dataRecordsets = await readDataRecordsets(pkg, documentPart, checkTime);
	} catch (error) {
		if (error instanceof VisioPackageError && error.code === 'LIMIT_RUNTIME') throw error;
		report(
			'unsupported-data-recordsets',
			`Saved external data could not be read and is preserved unchanged: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
	const stencils = only ? [] : await readVisioStencilWindows(pkg, documentPart);
	report(
		'cached-values-only',
		'ShapeSheet formulas, automatic connector routing, and external data are not evaluated; saved cached values and supported theme records are used.',
		{ severity: 'info' },
	);
	const snapGlue = readVisioSnapGlue(documentRoot);
	return {
		format: 'vsdx',
		...(snapGlue ? { snapGlue } : {}),
		pages,
		...(documentMasters.length ? { masters: documentMasters } : {}),
		...(stencils.length ? { stencils } : {}),
		diagnostics: diagnosticState.finish(),
		...(dataRecordsets.length ? { dataRecordsets } : {}),
		fontFamilies: [...new Set(resources.fonts.values())],
		...(comments.length ? { comments } : {}),
	};
}

/** Backgrounds first, selected page last. Missing pages and cycles are safe. */
export function getVisioPageLayers(document: VisioDocument, pageId: string): VisioPage[] {
	const result: VisioPage[] = [],
		seen = new Set<string>();
	const pages = new Map(document.pages.map((page) => [page.id, page]));
	let current = pages.get(pageId);
	while (current && !seen.has(current.id)) {
		seen.add(current.id);
		result.push(current);
		current =
			current.backgroundPageId === undefined ? undefined : pages.get(current.backgroundPageId);
	}
	return result.reverse();
}
