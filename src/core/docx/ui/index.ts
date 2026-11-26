export * from './backstage-focus';
export * from './break-note-schema';
export * from './collaboration-authority';
export * from './collaboration-identity';
export * from './collaboration-protocol';
export * from './collaboration';
export * from './comment-commands';
export { commentPointNodeSpec } from './comment-point';
export * from './comment-anchors';
export * from './comment-history';
export * from './yjs-collaboration';
export * from './yjs-media';
export * from './yjs-comments';
export * from './yjs-presence';
export * from './defined';
export * from './dirty-state';
export * from './field-guard';
export * from './table-structure-guard';
export * from '../table-grid';
export * from '../table-structure';
export { DEFAULT_TABLE_LOOK } from '../table-defaults';
export * from '../table-structure-merge';
export type { IdFactory } from '../table-structure-cells';
export * from './field-results';
export { emptySimpleFieldResult } from './simple-field-input';
export * from './header-footer-history';
export * from './hard-break-command';
export * from './link-commands';
export * from './run-format-command';
export * from './note-parts';
export * from './inline-content-schema';
export { chartNodeSpec, nodeChart, svgMarkupSpec } from './chart-spec';
export * from './chart-commands';
export * from './ligature-style';
export * from './node-numbering';
export * from './paragraph-attributes';
export * from './review-schema';
export * from './review-settings';
export * from './command-history';
export * from './review-commands';
export { formattingRevision } from './review-formatting';
export { paragraphFormattingRevision } from './review-paragraph-formatting';
export * from './track-changes-mode';
export * from './toggle-format';
export * from './customization';
export * from './ribbon-action-ids';
export * from './outline';
export * from './ribbon-colors';
export * from './ribbon-legacy-labels';
export * from './run-compare';
export * from './run-extra-mark';
export * from './run-marks';
export * from './run-adapter';
export * from './run-mark-properties';
export * from './schema-marks';
export * from './smartart-theme';
export * from './smartart-approximate';
export * from './table-render';
export * from './table-cell-views';
export * from './table-visuals';
export * from './theme/css-vars';
export * from './theme/defaults';
export * from './theme/element';
export * from './theme/types';
export * from './word-count';
export * from './zoom-fit';
export * from './content-control-commands';
export * from './content-control-guard';
export { controlOccurrences } from './content-control-occurrences';
export * from './content-control-spans';
export * from './symbol-font';
export * from './positional-tab-input';
export * from './form-field-commands';
export * from './form-field-guard';
export * from '../find-syntax';
export * from '../find-pattern';
export * from '../find-replacement';
export * from '../find-stories';
export * from '../find-format-criteria';
export * from '../find-format';
export {
	replaceFormattedMatches,
	type ReplaceFormatOptions,
	type ReplaceFormatResult,
} from '../find-format-replace';
export * from '../find-format-document';
export * from './edit-restriction';
export { EQUATION_GALLERY } from '../../math/equation-gallery';
/** Load table editing tools and gallery presets when a table command needs them. */
export const loadWordTableTools = () => import('./table-tools');
export type WordTableTools = Awaited<ReturnType<typeof loadWordTableTools>>;
export type { TableAutoFit, TableSort, TableTextSeparator } from './table-tools';
export const loadWordReferences = () => import('../references');
export type WordReferences = Awaited<ReturnType<typeof loadWordReferences>>;
export type {
	Bibliography,
	BibliographyAuthor,
	BibliographySource,
	BibliographyStyle,
	CitationOptions,
	IndexLeader,
	IndexOptions,
	ReferencePage,
	ReferenceTable,
	SourceField,
} from '../references';
