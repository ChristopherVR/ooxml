import { parseCoreProperties } from './core-properties';
import { readPackageInformation } from './package-information';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { declaredEntrySize, openZip } from '../opc/zip-limits';
import { inspectPackageSignatures, verifySignedPackage } from '../opc/signature/index';
import { resolveDocxLimits } from './package-limits';
import type { ZipLimits } from '../opc/zip-limits';
import type { Block, DocumentModel, DocxLoadOptions, LoadedDocument } from './model';
import { parseTwips } from './simple-types';
import { parseFontTable } from './font-table';
import { children, first, getW, parseXml, type XmlDocument } from './xml';
import { remember, saveDocx } from './save';
import { parseParagraphStyleCatalog } from './paragraph-styles';
import { parseBlocksFromContainer } from './block-parser';
import { parseDocumentParts } from './document-parts';
import { parseNumberingCatalog } from './numbering-parse';
import { parseComments } from './comments';
import { readCommentParts } from './comment-parts';
import { parseFormSettings, parseReviewPreferences, parseTrackChangesSetting } from './settings';
import { parsePageBackground } from './page-background';
import { parseRunStyleCatalog } from './character-styles';
import { readLatentStyles } from './latent-styles';
import { parseTableStyleCatalog } from './table-styles';
import { parseTheme, parseColorSchemeMapping } from './theme';

const px = (twips: string | undefined, fallback: number): number =>
	twips === undefined ? fallback : (Number(twips) * 96) / 1440;

import { type DrawingContext } from './drawing';
import {
	documentPartPath,
	parseContentTypes,
	parseRelationships,
	resolveInternalTarget,
} from './package-parts';
import { withParseWarnings } from './parse-diagnostics';
import {
	forEachParagraph,
	imageAndBookmarkWarnings,
	indexWordElements,
	warningsFor,
} from './parse-warnings';
import { diagramWarnings, resolveDocumentDiagrams } from './diagram-document';
import { documentDrawingBlocks, resolveDocumentCharts } from './chart';
import { readPlaceholderTexts } from './glossary';
import { readPackageProperties } from './package-properties';

export interface PackageContext {
	original: Uint8Array;
	/** The limits the package was opened under, which its saves reopen it with. */
	zipLimits: ZipLimits;
	sourceXml: string;
	base: DocumentModel;
}

/** The load warning for an enforced editing restriction (`w:documentProtection`), if any. */
function protectionWarning(edit: string): string | undefined {
	if (edit === 'forms')
		return 'The document is protected for filling in forms: the editor lets only form fields, content controls and sections left unprotected change. The protection password is checked only when you stop the protection.';
	if (edit === 'readOnly')
		return 'The document is protected against changes: the editor lets only its editable regions change. Stop protection (Review > Restrict Editing) to edit the rest.';
	if (edit === 'comments')
		return 'The document is protected for comments: the editor lets you add comments and change its editable regions. Stop protection (Review > Restrict Editing) to edit the rest.';
	if (edit === 'trackedChanges')
		return 'The document is protected for tracked changes: every change is recorded and cannot be accepted or rejected until the protection is stopped (Review > Restrict Editing).';
	return undefined;
}

const sizeText = (bytes: number): string =>
	bytes % (1024 * 1024) === 0 ? `${bytes / (1024 * 1024)} MiB` : `${bytes} bytes`;

export async function readPackage(
	input: Uint8Array | ArrayBuffer,
	options: DocxLoadOptions = {},
): Promise<{
	model: DocumentModel;
	context: PackageContext;
	media: ReadonlyMap<string, Uint8Array>;
}> {
	const limits = resolveDocxLimits(options.limits);
	const original =
		input instanceof Uint8Array ? new Uint8Array(input) : new Uint8Array(input.slice(0));
	const zip = await openZip(original, limits.zip);
	const file = zip.file('word/document.xml');
	if (!file) throw new Error('DOCX package has no word/document.xml part');
	const tooLarge = `word/document.xml exceeds the ${sizeText(limits.maxDocumentPartBytes)} limit (limits.maxDocumentPartBytes)`;
	if (declaredEntrySize(file) > limits.maxDocumentPartBytes) throw new Error(tooLarge);
	const sourceXml = await file.async('string');
	if (sourceXml.length > limits.maxDocumentPartBytes) throw new Error(tooLarge);
	const document: XmlDocument = parseXml(sourceXml);
	const scanned = indexWordElements(document);
	const body = scanned.get('body')?.[0];
	if (!body) throw new Error('DOCX document.xml has no w:body');
	const relsFile = zip.file('word/_rels/document.xml.rels');
	const contentTypesFile = zip.file('[Content_Types].xml');
	const mediaParts = new Set(Object.keys(zip.files).filter((name) => !zip.files[name]!.dir));
	const drawings: DrawingContext = {
		rels: parseRelationships(relsFile ? await relsFile.async('string') : undefined),
		contentTypes: parseContentTypes(
			contentTypesFile ? await contentTypesFile.async('string') : undefined,
		),
		mediaParts,
	};
	const schemaWarnings: string[] = [];
	const blocks: Block[] = withParseWarnings(schemaWarnings, () =>
		parseBlocksFromContainer(body, '', drawings),
	);
	const section = children(body, 'sectPr').at(-1);
	const size = first(section, 'pgSz');
	const margins = first(section, 'pgMar');
	const model: DocumentModel = {
		blocks,
		page: {
			width: px(getW(size, 'w'), 816),
			height: px(getW(size, 'h'), 1056),
			marginTop: px(getW(margins, 'top'), 96),
			marginRight: px(getW(margins, 'right'), 96),
			marginBottom: px(getW(margins, 'bottom'), 96),
			marginLeft: px(getW(margins, 'left'), 96),
		},
		warnings: warningsFor(scanned),
	};
	if ((await inspectPackageSignatures(zip)).hasSignatures)
		model.signatures = await verifySignedPackage(zip);
	const stylesFile = zip.file(documentPartPath(drawings.rels, 'styles', 'word/styles.xml'));
	if (stylesFile) {
		const stylesRoot = parseXml(await stylesFile.async('string')).documentElement;
		const catalogs = withParseWarnings(schemaWarnings, () => ({
			paragraph: parseParagraphStyleCatalog(stylesRoot),
			character: parseRunStyleCatalog(stylesRoot),
			table: parseTableStyleCatalog(stylesRoot),
		}));
		model.paragraphStyles = catalogs.paragraph;
		model.characterStyles = catalogs.character;
		model.tableStyles = catalogs.table;
		const latent = readLatentStyles(stylesRoot);
		if (latent) model.latentStyles = latent;
		model.warnings.push(
			'Paragraph and character styles resolve supported formatting through basedOn chains and document defaults. Style rendering and numbering-derived formatting can still differ from Word.',
			...model.paragraphStyles.warnings,
			...model.characterStyles.warnings,
			...model.tableStyles.warnings,
		);
	} else if (contentTypesFile && zip.file('_rels/.rels')) {
		model.usesImplicitNormalStyle = true;
	}
	const fontTableRel = [...drawings.rels.values()].find(
		(rel) => rel.mode === 'Internal' && rel.type.endsWith('/fontTable'),
	);
	const fontTablePart = fontTableRel
		? resolveInternalTarget(fontTableRel.target)
		: 'word/fontTable.xml';
	const fontTable = zip.file(fontTablePart);
	if (fontTable) {
		const fallbacks = parseFontTable(await fontTable.async('string'));
		if (fallbacks.size) model.fontFallbacks = fallbacks;
	}
	const themeFile = zip.file(documentPartPath(drawings.rels, 'theme', 'word/theme/theme1.xml'));
	const settingsFile = zip.file('word/settings.xml');
	if (themeFile) {
		model.theme = parseTheme(await themeFile.async('string'));
		const themeSettings = zip.file(
			documentPartPath(drawings.rels, 'settings', 'word/settings.xml'),
		);
		if (themeSettings) {
			const settingsXml = await themeSettings.async('string');
			model.theme.colorMapping = withParseWarnings(schemaWarnings, () =>
				parseColorSchemeMapping(settingsXml),
			);
		}
		model.warnings.push(
			'Theme colors and fonts resolve for rendering through a separate layer; direct theme references are preserved and never flattened onto runs. Theme tint and shade quantization is approximate.',
		);
	}
	const numberingFile = zip.file('word/numbering.xml');
	if (numberingFile) {
		const numberingXml = await numberingFile.async('string');
		const numberingRels = zip.file('word/_rels/numbering.xml.rels');
		const numberingRelationships = parseRelationships(
			numberingRels ? await numberingRels.async('string') : undefined,
		);
		model.numberingCatalog = withParseWarnings(schemaWarnings, () =>
			parseNumberingCatalog(numberingXml, numberingRelationships),
		);
		model.warnings.push(
			'Numbering is resolved for decimal, Roman numeral, letter, ordinal, cardinal/ordinal text and bullet formats, including multilevel lvlText and legal numbering. Style-linked numbering and other custom formats fall back to Decimal Number and are not rendered as Word would.',
			...model.numberingCatalog.warnings,
		);
	} else if (scanned.has('numPr')) {
		model.warnings.push(
			'Paragraphs reference list numbering, but the package has no word/numbering.xml part; numbering is preserved as paragraph XML but not rendered.',
		);
	}
	if (blocks.some((block) => block.type === 'table'))
		model.warnings.push(
			'Table grid widths, merges, borders, shading, cell alignment/margins and table style conditional formatting resolve for rendering; row/column fragmentation and full table-style precedence are not modeled.',
		);
	const parts = await parseDocumentParts(zip, body, blocks, drawings);
	model.sections = parts.sections;
	if (parts.compatibilityMode !== undefined) model.compatibilityMode = parts.compatibilityMode;
	if (parts.doNotUseHtmlParagraphAutoSpacing) model.doNotUseHtmlParagraphAutoSpacing = true;
	if (parts.kinsoku) model.kinsoku = parts.kinsoku;
	if (parts.footnotes) model.footnotes = parts.footnotes;
	if (parts.footnoteSeparators) model.footnoteSeparators = parts.footnoteSeparators;
	if (parts.endnotes) model.endnotes = parts.endnotes;
	if (parts.footnoteNumFmt) model.footnoteNumFmt = parts.footnoteNumFmt;
	if (parts.endnoteNumFmt) model.endnoteNumFmt = parts.endnoteNumFmt;
	if (parts.evenAndOddHeaders) model.evenAndOddHeaders = true;
	if (parts.autoHyphenation) model.autoHyphenation = true;
	if (parts.hyphenationZoneTwips !== undefined)
		model.hyphenationZoneTwips = parts.hyphenationZoneTwips;
	if (parts.consecutiveHyphenLimit !== undefined)
		model.consecutiveHyphenLimit = parts.consecutiveHyphenLimit;
	if (parts.doNotHyphenateCaps) model.doNotHyphenateCaps = true;
	const coreFile = zip.file('docProps/core.xml');
	if (coreFile) {
		const properties = parseCoreProperties(await coreFile.async('string'));
		if (properties) model.properties = properties;
	}
	const packageProperties = await readPackageProperties(zip);
	if (packageProperties) model.packageProperties = packageProperties;
	const packageInformation = await readPackageInformation(zip);
	if (packageInformation) model.packageInformation = packageInformation;
	await readPlaceholderTexts(zip, drawings.rels, model);
	const pageColor = parsePageBackground(document.documentElement, model.theme);
	if (pageColor) model.pageColor = pageColor;
	model.warnings.push(...parts.warnings);
	if (
		blocks.some(
			(block) => block.type === 'table' && !block.structureEditable && !block.restructurable,
		)
	)
		model.warnings.push(
			'Merged, nested, or complex tables can be read, but their row and column structure cannot be edited safely.',
		);
	const commentParts = await readCommentParts(zip);
	if (commentParts.comments.xml) {
		model.comments = parseComments(
			commentParts.comments.xml,
			commentParts.extended.xml,
			commentParts.ids.xml,
			commentParts.extensible.xml,
		);
		model.warnings.push(
			'Comment ranges are modeled at run granularity (they may span paragraphs); a range edge outside any run, such as around an empty paragraph or a table, moves to the nearest commented text when that paragraph is edited.',
		);
	}
	if (settingsFile) {
		const settingsXml = await settingsFile.async('string');
		model.trackChanges = parseTrackChangesSetting(settingsXml);
		const tabStop = parseTwips(
			getW(first(parseXml(settingsXml).documentElement, 'defaultTabStop'), 'val'),
		);
		if (tabStop && tabStop > 0) model.defaultTabStopTwips = tabStop;
		Object.assign(model, parseReviewPreferences(settingsXml), parseFormSettings(settingsXml));
		const warning = model.protection?.enforced && protectionWarning(model.protection.edit);
		if (warning) model.warnings.push(warning);
	}
	if (
		blocks.some(
			(block) =>
				block.type === 'table' &&
				block.rows.some((row) => row.some((cell) => cell.nestedTables?.length)),
		)
	)
		model.warnings.push(
			'Text and structure (rows, columns, merged and split cells) of nested tables are editable; their borders, shading and other formatting are preserved but cannot be changed.',
		);
	// Every block that can hold a drawing: body, headers, footers, footnotes and endnotes.
	const pictureBlocks = documentDrawingBlocks({ ...model, blocks });
	await resolveDocumentDiagrams(pictureBlocks, async (partName) =>
		zip.file(partName)?.async('string'),
	);
	await resolveDocumentCharts(
		pictureBlocks,
		async (partName) => zip.file(partName)?.async('string'),
		async (partName) => zip.file(partName)?.async('uint8array'),
	);
	model.warnings.push(...imageAndBookmarkWarnings(blocks), ...diagramWarnings(pictureBlocks));
	for (const warning of schemaWarnings)
		if (!model.warnings.includes(warning)) model.warnings.push(warning);
	// Pictures anywhere in the document: body, headers, footers, footnotes and endnotes.
	const imagePartNames = new Set<string>();
	forEachParagraph(pictureBlocks, (paragraph) => {
		for (const run of paragraph.runs)
			for (const part of [
				run.image?.partName,
				run.image?.svgPartName,
				...Object.values(run.image?.diagram?.drawing?.pictureParts ?? {}),
				...Object.values(run.image?.diagram?.layoutSource?.pictureParts ?? {}),
			])
				if (part) imagePartNames.add(part);
	});
	for (const bullet of Object.values(model.numberingCatalog?.pictureBullets ?? {}))
		imagePartNames.add(bullet.partName);
	const media = new Map<string, Uint8Array>();
	for (const partName of imagePartNames) {
		const part = zip.file(partName);
		if (part) media.set(partName, await part.async('uint8array'));
	}
	const context = { original, zipLimits: limits.zip, sourceXml, base: structuredClone(model) };
	remember(model, context);
	return { model, context, media };
}

export async function loadDocx(
	input: Uint8Array | ArrayBuffer,
	options: DocxLoadOptions = {},
): Promise<LoadedDocument> {
	const { model, context, media } = await readPackage(input, options);
	return {
		model,
		media,
		save: (next = model, pendingMedia, options) => {
			if (next !== model) remember(next, context);
			return saveDocx(next, pendingMedia, options);
		},
	};
}
