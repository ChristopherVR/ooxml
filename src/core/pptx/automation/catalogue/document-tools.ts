import * as schemas from '../schemas/index';
import * as customShowTools from '../tools/custom-show-tools';
import * as exportTools from '../tools/export-tools';
import * as geometryTools from '../tools/geometry-tools';
import * as hyperlinkTools from '../tools/hyperlink-tools';
import * as jsonTools from '../tools/json-tools';
import * as layoutTools from '../tools/layout-tools';
import * as lockTools from '../tools/lock-tools';
import * as metadataTools from '../tools/metadata-tools';
import * as oleTools from '../tools/ole-tools';
import * as presentationTools from '../tools/presentation-tools';
import * as sectionTools from '../tools/section-tools';
import type { AutomationTool } from '../../../automation/tools/types';
import { deckTool } from './deck-tool';

/** PowerPoint tools: metadata, sections, export, JSON, hyperlinks, geometry, locks, properties, layouts and OLE objects. */
export const documentTools: readonly AutomationTool[] = [
	deckTool(
		'read',
		'get_metadata',
		'Get presentation metadata (title, author, keywords, custom properties)',
		schemas.GetMetadataSchema,
		(ctx) => metadataTools.getMetadata(ctx),
	),
	deckTool(
		'edit',
		'update_metadata',
		'Update presentation metadata (title, author, company, custom properties)',
		schemas.UpdateMetadataSchema,
		(ctx, params) => metadataTools.updateMetadata(ctx, params),
	),
	deckTool(
		'edit',
		'manage_sections',
		'List, add, remove, reorder sections or move slides between sections',
		schemas.ManageSectionsSchema,
		(ctx, params) => sectionTools.manageSections(ctx, params),
		{ returnsText: true },
	),
	deckTool(
		'edit',
		'manage_custom_shows',
		'List, create, rename, copy or delete custom shows, or edit their slide lists (add, remove, reorder)',
		schemas.ManageCustomShowsSchema,
		(ctx, params) => customShowTools.manageCustomShows(ctx, params),
		{ returnsText: true },
	),
	deckTool(
		'read',
		'export_to_svg',
		'Export all or specific slides to SVG strings',
		schemas.ExportToSvgSchema,
		(ctx, params) => exportTools.exportToSvg(ctx, params),
	),
	deckTool(
		'read',
		'export_slide_svg',
		'Export a single slide to SVG',
		schemas.ExportSlideSvgSchema,
		(ctx, params) => exportTools.exportSlideSvg(ctx, params),
	),
	deckTool(
		'read',
		'export_to_json',
		'Export the presentation to the portable pptx-viewer-json document format ' +
			'(self-contained, binary assets embedded as base64)',
		schemas.ExportToJsonSchema,
		(ctx, params) => jsonTools.exportToJson(ctx, params),
	),
	deckTool(
		'edit',
		'import_from_json',
		'Replace the presentation content with a deck imported from a pptx-viewer-json document',
		schemas.ImportFromJsonSchema,
		(ctx, params) => jsonTools.importFromJson(ctx, params),
	),
	deckTool(
		'edit',
		'manage_hyperlinks',
		'List, set, or remove hyperlinks/actions on elements',
		schemas.ManageHyperlinksSchema,
		(ctx, params) => hyperlinkTools.manageHyperlinks(ctx, params),
		{ returnsText: true },
	),
	deckTool(
		'edit',
		'replace_geometry',
		'Replace a shape element geometry with a preset or custom SVG path',
		schemas.ReplaceGeometrySchema,
		(ctx, params) =>
			geometryTools.replaceGeometry(ctx, {
				...params,
				adjustments: params.adjustments as Record<string, number> | undefined,
			}),
	),
	deckTool(
		'edit',
		'set_element_lock',
		'Lock or unlock an element (prevent move, resize, rotation, selection, text edit)',
		schemas.SetElementLockSchema,
		(ctx, params) => lockTools.setElementLockT(ctx, params),
	),
	deckTool(
		'read',
		'get_presentation_properties',
		'Get slideshow properties (show type, loop, advance mode, slide range)',
		schemas.GetPresentationPropertiesSchema,
		(ctx) => presentationTools.getPresentationProperties(ctx),
	),
	deckTool(
		'edit',
		'update_presentation_properties',
		'Update slideshow properties (show type, loop, advance mode, pen color)',
		schemas.UpdatePresentationPropertiesSchema,
		(ctx, params) => presentationTools.updatePresentationProperties(ctx, params),
	),
	deckTool(
		'read',
		'get_layouts',
		'List available slide layouts from the presentation masters',
		schemas.GetLayoutsSchema,
		(ctx) => layoutTools.getLayouts(ctx),
	),
	deckTool(
		'edit',
		'apply_layout',
		'Apply a slide layout to a specific slide',
		schemas.ApplyLayoutSchema,
		(ctx, params) => layoutTools.applyLayout(ctx, params),
	),
	deckTool(
		'read',
		'ole_get_content',
		"Inspect an embedded OLE object's editable content: resolved kind, plus the sheet grid, " +
			'paragraph list, or nested-deck slide summaries, whichever applies',
		schemas.GetOleContentSchema,
		(ctx, params) => oleTools.getOleContent(ctx, params),
	),
	deckTool(
		'edit',
		'ole_set_sheet_cell',
		"Edit one cell of an embedded Excel OLE object's first worksheet",
		schemas.SetOleSheetCellSchema,
		(ctx, params) => oleTools.setOleSheetCell(ctx, params),
	),
	deckTool(
		'edit',
		'ole_set_document_paragraph',
		"Replace one paragraph's text in an embedded Word (.docx or legacy .doc) OLE object",
		schemas.SetOleDocumentParagraphSchema,
		(ctx, params) => oleTools.setOleDocumentParagraph(ctx, params),
	),
	deckTool(
		'edit',
		'ole_set_deck_slide_title',
		'Edit one text-bearing shape on one slide of an embedded nested-deck (PowerPoint) OLE ' +
			'object, via a full load/save round-trip through the nested deck (call ole_get_content ' +
			"first to discover each slide's elementIds)",
		schemas.SetOleDeckSlideTitleSchema,
		(ctx, params) => oleTools.setOleDeckSlideTitle(ctx, params),
	),
	deckTool(
		'edit',
		'ole_replace_file',
		"Replace an embedded OLE object's payload wholesale with an arbitrary file; works for any " +
			'OLE kind, including a generic "Package" object or an unsupported format',
		schemas.ReplaceOleFileSchema,
		(ctx, params) => oleTools.replaceOleFileT(ctx, params),
	),
	deckTool(
		'edit',
		'ole_set_object_name',
		"Rename an OLE object's Object Name (p:oleObj/@name); regenerates the baked-in icon " +
			'caption too when the object is displayed as an icon',
		schemas.SetOleObjectNameSchema,
		(ctx, params) => oleTools.setOleObjectNameT(ctx, params),
	),
];
