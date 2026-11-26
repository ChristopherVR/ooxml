export {
	REFERENCE_LCID,
	authorityEntryCode,
	indexEntryCode,
	referenceFieldRuns,
} from './instructions';
export { authorityCode } from './authorities';
export { indexCode, INDEX_LEADERS, type IndexLeader, type IndexOptions } from './index-entries';
export { indexSection } from './index-section';
export { type ReferencePage } from './entries';
export { buildReferenceTable, findReferenceTables, type ReferenceTable } from './tables';
export {
	citationCode,
	citationEvaluator,
	citationSpans,
	editCitationCode,
	readCitation,
	refreshCitations,
	type CitationOptions,
} from './citations';
export { readBibliographyXml, writeBibliographyXml } from './source-xml';
export * from './source-model';
