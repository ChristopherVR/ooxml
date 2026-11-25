import { applyStyleCatalogs } from './style-catalog-save';
import { applyWordTheme } from './theme-save';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import JSZip from 'jszip';
import { openZip } from '../opc/zip-limits';
import { dropAddedFolders, entryNames } from '../opc/zip-folders';
import { stripPackageSignatures } from '../opc/signature/index';
import type { DocumentModel, PendingMediaPart, SaveOptions } from './model';
import { stampModified } from './save-modified';
import { applyDataBindings } from './data-binding';
import { buildXml, parseXml, type XmlDocument } from './xml';
import type { PackageContext } from './parse';
import { applyModel } from './write';
import { DocPrIdAllocator } from './docpr-ids';
import { assertValidDocumentModel } from './validate-model';
import { applyCoreProperties } from './core-properties';
import { applyPackageInformation } from './package-information-save';
import { scrubPersonalParts } from './personal-information';
import { applyNumberingCatalog } from './numbering-package';
import { applyHeaderFooterEdits, createHeaderFooterParts } from './write-header-footer';
import { applyNoteEdits } from './write-notes';
import { writeNewRelationships } from './part-relationships';
import {
	applyCompatibilityMode,
	applyDocumentProtection,
	applySettingsFlag,
	applyTrackChangesSetting,
} from './settings';
import { applyPageBackground } from './page-background';
import { applyComments } from './write-comments';
import { numberCommentIds } from './comment-spans';
import { refreshStyleRefFields } from './style-ref-fields';
import { refreshComputedFields } from './computed-fields';
import { applyDiagramEdits } from './diagram-save';
import { applyGeneratedDiagramCaches } from './diagram-cache';
import { applyChartPartEdits } from './chart-save';

import { maxWordId, numberRevisionIds } from './revision-ids';
import { numberMoveNames } from './move-names';
import { parseRelationships } from './package-parts';

const RELS_PART = 'word/_rels/document.xml.rels';

const contexts = new WeakMap<DocumentModel, PackageContext>();

/**
 * Saves of `model` patch the package `context` was read from. Saving never changes
 * `context.base` (the model as loaded), so every model saved against one package shares it.
 */
export function remember(model: DocumentModel, context: PackageContext): void {
	contexts.set(model, context);
}
function newDocument(): XmlDocument {
	return parseXml(
		'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
	);
}

/**
 * `model` with its header and footer STYLEREF and computed field results (DOCPROPERTY...)
 * brought up to date, or undefined when none changes. Word updates those fields as it draws each
 * page, so a saved header should not keep a stale result; a part the writer cannot rebuild with
 * the new results keeps the ones it has, and a model the update cannot read is saved as it is:
 * the update never makes a save fail.
 */
/** `model` with its body's computed field results up to date, or undefined when none changes. */
function bodyResults(model: DocumentModel): DocumentModel | undefined {
	try {
		const refreshed = refreshComputedFields(model, { body: true });
		if (refreshed.blocks === model.blocks) return undefined;
		const body = { ...model, blocks: refreshed.blocks };
		assertValidDocumentModel(body);
		return body;
	} catch {
		return undefined;
	}
}

function headerFooterResults(model: DocumentModel): DocumentModel | undefined {
	try {
		const refreshed = refreshStyleRefFields(refreshComputedFields(model));
		if (refreshed === model) return undefined;
		assertValidDocumentModel(refreshed);
		return refreshed;
	} catch {
		return undefined;
	}
}

export async function saveDocx(
	model: DocumentModel,
	pendingMedia?: ReadonlyMap<string, PendingMediaPart>,
	options?: SaveOptions,
): Promise<Uint8Array> {
	const ids = new Set<string>();
	for (const block of model.blocks) {
		if (ids.has(block.id)) throw new Error(`Duplicate document block id: ${block.id}`);
		ids.add(block.id);
		if (block.type === 'table')
			for (const row of block.rows)
				for (const cell of row)
					for (const paragraph of cell.paragraphs) {
						if (ids.has(paragraph.id))
							throw new Error(`Duplicate document node id: ${paragraph.id}`);
						ids.add(paragraph.id);
					}
	}
	const binding = contexts.get(model);
	if (binding && JSON.stringify(model) === JSON.stringify(binding.base))
		return new Uint8Array(binding.original);
	assertValidDocumentModel(model);
	model = numberMoveNames(
		numberRevisionIds(numberCommentIds(model), maxWordId(binding?.sourceXml)),
	);
	const zip = binding ? await openZip(binding.original, binding.zipLimits) : new JSZip();
	const sourceEntries = entryNames(zip);
	let document = binding ? parseXml(binding.sourceXml) : newDocument();
	const docPrIds = new DocPrIdAllocator();
	await docPrIds.reserveFromPackage(zip);
	if (!binding) {
		zip.file(
			'[Content_Types].xml',
			'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
		);
		zip.file(
			'_rels/.rels',
			'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
		);
	}
	// Header/footer parts the model added exist (with their relationships) before the body is written.
	const headerFooterIds = await createHeaderFooterParts(zip, model, binding?.base, docPrIds);
	const existingRelationships = parseRelationships(await zip.file(RELS_PART)?.async('string'));
	const write = (target: XmlDocument, source: DocumentModel) =>
		applyModel(
			target,
			source,
			binding?.base.blocks ?? [],
			existingRelationships.keys(),
			binding?.base.sections ?? [],
			docPrIds,
			headerFooterIds,
		);
	// The body's computed field results (DOCPROPERTY...) are saved up to date; a body the writer
	// cannot rebuild with them is written as the model holds it.
	const preferred = bodyResults(model);
	let written: ReturnType<typeof write> | undefined;
	if (preferred)
		try {
			written = write(document, preferred);
		} catch {
			document = binding ? parseXml(binding.sourceXml) : newDocument();
		}
	const { newRelationships } = written ?? write(document, model);
	if (model.pageColor !== binding?.base.pageColor) applyPageBackground(document, model.pageColor);
	zip.file('word/document.xml', buildXml(document));
	await writeNewRelationships(zip, 'word/document.xml', newRelationships, pendingMedia);
	if (binding) await applyDiagramEdits(zip, model.blocks, binding.base.blocks, pendingMedia);
	await applyGeneratedDiagramCaches(zip, model.blocks, binding?.base.blocks ?? [], model.theme);
	await applyHeaderFooterEdits(
		zip,
		model,
		binding?.base ?? { ...model, sections: [] },
		pendingMedia,
		docPrIds,
		headerFooterResults(model),
	);
	// A new document starts without notes, so every note is created along with its part.
	await applyNoteEdits(
		zip,
		model,
		binding?.base ?? { ...model, footnotes: [], endnotes: [] },
		pendingMedia,
		docPrIds,
	);
	// Edited charts are written back with their workbooks; unedited chart parts are not touched.
	await applyChartPartEdits(zip, model, binding?.base);
	await applyNumberingCatalog(zip, model, binding, pendingMedia);
	await applyCoreProperties(zip, model.properties, binding?.base.properties);
	await applyStyleCatalogs(zip, model, binding?.base);
	await applyWordTheme(zip, model.theme, binding?.base.theme);
	if (model.compatibilityMode !== binding?.base.compatibilityMode)
		await applyCompatibilityMode(zip, model.compatibilityMode);
	if (Boolean(model.evenAndOddHeaders) !== Boolean(binding?.base.evenAndOddHeaders))
		await applySettingsFlag(zip, 'evenAndOddHeaders', Boolean(model.evenAndOddHeaders));
	// Word shows a page colour only when settings.xml asks for background shapes.
	if (model.pageColor !== binding?.base.pageColor)
		await applySettingsFlag(zip, 'displayBackgroundShape', Boolean(model.pageColor));
	if (Boolean(model.autoHyphenation) !== Boolean(binding?.base.autoHyphenation))
		await applySettingsFlag(zip, 'autoHyphenation', Boolean(model.autoHyphenation));
	if (model.trackChanges !== binding?.base.trackChanges)
		await applyTrackChangesSetting(zip, model.trackChanges === true);
	if (JSON.stringify(model.protection) !== JSON.stringify(binding?.base.protection))
		await applyDocumentProtection(zip, model.protection);
	for (const [key, flag] of [
		['trackFormatting', 'doNotTrackFormatting'],
		['trackMoves', 'doNotTrackMoves'],
	] as const)
		if ((model[key] !== false) !== (binding?.base[key] !== false))
			await applySettingsFlag(zip, flag, model[key] === false);
	if (JSON.stringify(model.comments ?? []) !== JSON.stringify(binding?.base.comments ?? []))
		await applyComments(zip, model.comments ?? []);
	if (binding) await applyDataBindings(zip, model, binding.base, binding.sourceXml);
	await applyPackageInformation(zip, model.packageInformation, binding?.base.packageInformation);
	const privacy = model.packageInformation;
	if (privacy?.removePersonalInformation)
		await scrubPersonalParts(zip, privacy.removeDateAndTime === true);
	if (binding)
		await stampModified(zip, options?.modified, privacy?.removePersonalInformation === true);
	if (binding) await stripPackageSignatures(zip);
	dropAddedFolders(zip, sourceEntries);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
