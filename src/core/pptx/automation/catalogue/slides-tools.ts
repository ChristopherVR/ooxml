import * as schemas from '../schemas/index';
import * as contentTools from '../tools/content-tools';
import * as elementTools from '../tools/element-tools';
import * as slideTools from '../tools/slide-tools';
import * as styleTools from '../tools/style-tools';
import * as tableStyleTools from '../tools/table-style-tools';
import * as tableTools from '../tools/table-tools';
import type { AutomationTool } from '../../../automation/tools/types';
import { deckTool } from './deck-tool';

/** PowerPoint tools: slides, elements, tables, styles and text. */
export const slidesTools: readonly AutomationTool[] = [
	deckTool(
		'read',
		'get_slide',
		'Inspect a single slide: returns layout, notes, elements.',
		schemas.GetSlideSchema,
		(ctx, params) => slideTools.getSlide(ctx, { slideIndex: params.slideIndex }),
	),
	deckTool(
		'edit',
		'add_slide',
		'Add a new blank slide to the presentation',
		schemas.AddSlideSchema,
		(ctx, params) => slideTools.addSlide(ctx, params),
	),
	deckTool(
		'edit',
		'delete_slides',
		'Delete one or more slides by index',
		schemas.DeleteSlidesSchema,
		(ctx, params) => slideTools.deleteSlides(ctx, { slideIndexes: params.slideIndexes }),
	),
	deckTool(
		'edit',
		'reorder_slides',
		'Reorder slides by providing a new index array',
		schemas.ReorderSlidesSchema,
		(ctx, params) => slideTools.reorderSlides(ctx, { newOrder: params.newOrder }),
	),
	deckTool(
		'edit',
		'duplicate_slide',
		'Duplicate a slide with new element IDs',
		schemas.DuplicateSlideSchema,
		(ctx, params) => slideTools.duplicateSlide(ctx, params),
	),
	deckTool(
		'edit',
		'update_slide_properties',
		'Update slide background, notes, or visibility',
		schemas.UpdateSlidePropertiesSchema,
		(ctx, params) => slideTools.updateSlideProperties(ctx, params),
	),
	deckTool(
		'edit',
		'set_slide_transition',
		'Set or remove a slide transition effect',
		schemas.SetSlideTransitionSchema,
		(ctx, params) => slideTools.setSlideTransition(ctx, params),
	),
	deckTool(
		'edit',
		'set_canvas_size',
		'Change the presentation canvas dimensions',
		schemas.SetCanvasSizeSchema,
		(ctx, params) => slideTools.setCanvasSize(ctx, params),
	),
	deckTool(
		'edit',
		'add_element',
		'Add a text, shape, image, table, or connector element',
		schemas.AddElementSchema,
		(ctx, params) => elementTools.addElement(ctx, params),
	),
	deckTool(
		'edit',
		'update_element',
		'Update element position, size, text, or style',
		schemas.UpdateElementSchema,
		(ctx, params) => elementTools.updateElement(ctx, params),
	),
	deckTool(
		'edit',
		'rename_element',
		"Rename an element (the selection pane's display name; cNvPr/@name)",
		schemas.RenameElementSchema,
		(ctx, params) => elementTools.renameElement(ctx, params),
	),
	deckTool(
		'edit',
		'delete_elements',
		'Delete one or more elements by ID',
		schemas.DeleteElementsSchema,
		(ctx, params) => elementTools.deleteElements(ctx, params),
	),
	deckTool(
		'edit',
		'arrange_elements',
		'Align elements or change z-order (layer)',
		schemas.ArrangeElementsSchema,
		(ctx, params) => elementTools.arrangeElements(ctx, params),
	),
	deckTool(
		'edit',
		'clone_element',
		'Clone an element within or across slides',
		schemas.CloneElementSchema,
		(ctx, params) => elementTools.cloneElement(ctx, params),
	),
	deckTool(
		'edit',
		'set_element_animation',
		'Set entrance/exit animation on an element',
		schemas.SetElementAnimationSchema,
		(ctx, params) => elementTools.setElementAnimation(ctx, params),
	),
	deckTool(
		'edit',
		'group_elements',
		'Group multiple elements into a group',
		schemas.GroupElementsSchema,
		(ctx, params) => elementTools.groupElements(ctx, params),
	),
	deckTool(
		'edit',
		'ungroup_elements',
		'Ungroup a group element back to individual elements',
		schemas.UngroupElementsSchema,
		(ctx, params) => elementTools.ungroupElements(ctx, params),
	),
	deckTool(
		'edit',
		'batch_update_elements',
		'Apply the same properties to multiple elements',
		schemas.BatchUpdateElementsSchema,
		(ctx, params) => elementTools.batchUpdateElements(ctx, params),
	),
	deckTool(
		'edit',
		'update_table_cells',
		'Update text content of table cells',
		schemas.UpdateTableCellsSchema,
		(ctx, params) => tableTools.updateTableCells(ctx, params),
	),
	deckTool(
		'edit',
		'manage_table_structure',
		'Insert or delete table rows and columns',
		schemas.ManageTableStructureSchema,
		(ctx, params) => tableTools.manageTableStructure(ctx, params),
	),
	deckTool(
		'edit',
		'set_table_style_section',
		'Patch the fill, text, borders, or cell3D of one section (wholeTbl, banding, first/last row/col, or a corner cell) of an existing loaded table style',
		schemas.SetTableStyleSectionSchema,
		(ctx, params) => tableStyleTools.setTableStyleSection(ctx, params),
	),
	deckTool(
		'edit',
		'create_table_style',
		'Create a new table style (optionally cloned from an existing one) and optionally set it as the default',
		schemas.CreateTableStyleSchema,
		(ctx, params) => tableStyleTools.createTableStyle(ctx, params),
	),
	deckTool(
		'edit',
		'delete_table_style',
		'Remove a table style from the presentation',
		schemas.DeleteTableStyleSchema,
		(ctx, params) => tableStyleTools.deleteTableStyle(ctx, params),
	),
	deckTool(
		'edit',
		'assign_table_style',
		'Assign a table style (and row/column emphasis flags) to an existing table',
		schemas.AssignTableStyleSchema,
		(ctx, params) => tableStyleTools.assignTableStyle(ctx, params),
	),
	deckTool(
		'edit',
		'update_element_style',
		'Update fill, stroke, shadow, glow, or image effects',
		schemas.UpdateElementStyleSchema,
		(ctx, params) =>
			styleTools.updateElementStyle(ctx, params as styleTools.UpdateElementStyleParams),
	),
	deckTool(
		'read',
		'run_accessibility_check',
		'Audit the presentation for accessibility issues',
		schemas.AccessibilityCheckSchema,
		(ctx) => styleTools.runAccessibilityCheck(ctx),
	),
	deckTool(
		'read',
		'find_text',
		'Search for text across all slides',
		schemas.FindTextSchema,
		(ctx, params) => contentTools.findText(ctx, params),
	),
	deckTool(
		'edit',
		'replace_text',
		'Find and replace text across slides',
		schemas.ReplaceTextSchema,
		(ctx, params) => contentTools.replaceText(ctx, params),
		{ returnsText: true },
	),
	deckTool(
		'edit',
		'manage_comments',
		'List, add, delete, or resolve slide comments',
		schemas.ManageCommentsSchema,
		(ctx, params) => contentTools.manageComments(ctx, params),
		{ returnsText: true },
	),
];
