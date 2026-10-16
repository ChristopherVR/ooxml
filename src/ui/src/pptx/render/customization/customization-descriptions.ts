/**
 * One-line English descriptions for the customisation ids that have no UI
 * label of their own (panels, features, dialogs, toolbar buttons, shortcut
 * commands). Used by the generated docs reference; the `Record` types make a
 * missing description a compile error.
 *
 * @module render/customization/customization-descriptions
 */
import type { EditorKeyActionName } from '../editor-keymap';
import type { ToolbarButtonId } from '../toolbar-actions';
import type { ViewerDialogRegionId, ViewerPanelRegionId } from './customization-regions';
import type {
	ViewerCapabilityId,
	ViewerDialogId,
	ViewerExportFormatId,
	ViewerFeatureId,
	ViewerPanelId,
} from './customization-types';

export const TOOLBAR_BUTTON_DESCRIPTIONS: Record<ToolbarButtonId, string> = {
	share: 'Share / collaboration button in the tab row and mobile toolbar.',
	broadcast: 'Broadcast (present online) button.',
	export: 'Export button and the File > Export page.',
	undo: 'Undo button in the quick-access strip.',
	redo: 'Redo button in the quick-access strip.',
	record: 'Record button and the Record ribbon tab.',
	notes: 'Notes toggle in the status bar.',
	fullscreen: 'Full-screen toggle.',
	zoom: 'The zoom cluster (zoom in, zoom out, fit).',
	navigation: 'The previous / next slide cluster.',
	mergeShapes:
		'The Merge Shapes dropdown (Union, Combine, Fragment, Intersect, Subtract) in the Home tab Arrange group.',
	crop: 'Both picture Crop buttons (crop mode, aspect ratio, Fill, Fit): Home > Arrange, Picture Format > Size.',
};

export const VIEWER_PANEL_DESCRIPTIONS: Record<ViewerPanelId, string> = {
	statusBar: 'The status bar under the canvas (slide counter, zoom, view buttons).',
	slidesPane: 'The slide thumbnail rail on the left.',
	inspector: 'The format / properties inspector on the right.',
	notes: 'The speaker-notes panel under the canvas.',
	quickAccessToolbar: 'The quick-access strip in the title bar (save, undo, redo, ...).',
	titleBar: 'The title bar above the ribbon (file name, quick access, account).',
	selectionPane: 'The Selection Pane and View > Selection Pane.',
	findReplace: 'The Find and Replace panel, Home > Find / Replace and their shortcuts.',
	accessibilityPane:
		'The accessibility checker panel, Review > Check Accessibility and Help > Accessibility checker.',
	comparePane: 'The Compare panel and Review > Compare.',
};

export const VIEWER_FEATURE_DESCRIPTIONS: Record<ViewerFeatureId, string> = {
	ai: 'The AI assistant: toolbar toggle, chat panel, AI context-menu entries, AI Options page.',
	collaboration: 'Real-time collaboration: Share and Broadcast buttons, dialogs and File pages.',
	comments:
		'Commenting: the Add Comment context-menu entry and the Comments groups of Insert and Review.',
	presentMode: 'Slide-show entry points: the Slide Show ribbon tab.',
	editPoints:
		'Edit Points: the Edit Points context-menu entry and the point-editing mode it opens.',
};

export const VIEWER_DIALOG_DESCRIPTIONS: Record<ViewerDialogId, string> = {
	options: 'File > Options (the Settings dialog) and the File tab entry that opens it.',
	share: 'The Share dialog, its toolbar button, File > Share page and card.',
	broadcast: 'The Broadcast dialog and its button.',
	print: 'The Print dialog, File > Print page and card.',
	export: 'File > Export page, its cards and the Export button.',
	hyperlink: 'The Insert Hyperlink dialog: Insert > Link, the Link menu entry and its shortcut.',
	equation: 'The Equation editor and Insert > Equation.',
	pasteSpecial: 'The Paste Special dialog and its shortcut.',
	slideTemplates:
		'The slide template gallery and its entry at the end of the Home > New Slide menu.',
	passwordProtection: 'The password protection dialog and the File > Info Protect card.',
	signatures: 'The digital signatures dialog and the File > Info Signatures card.',
};

export const VIEWER_CAPABILITY_DESCRIPTIONS: Record<ViewerCapabilityId, string> = {
	edit: 'Editing. `false` opens the deck read-only whatever the edit prop says.',
	export: 'Exporting: the same as hiding the `export` dialog.',
	print: 'Printing: the same as hiding the `print` dialog.',
	share: 'Sharing: hides the `share` and `broadcast` dialogs.',
	copy: 'Copying content out: Copy, Cut, Save as Picture, their shortcuts and Copy as image.',
	present: 'Presenting: the same as disabling the `presentMode` feature.',
	comment: 'Commenting: the same as disabling the `comments` feature.',
};

export const VIEWER_PANEL_REGION_DESCRIPTIONS: Record<ViewerPanelRegionId, string> = {
	inspector: 'The right-hand inspector column (the pane, not its open/close state).',
	slidesPane: 'The slide thumbnail rail on the left.',
	notes: 'The speaker-notes panel under the canvas.',
	statusBar: 'The status bar under the canvas.',
};

export const VIEWER_DIALOG_REGION_DESCRIPTIONS: Record<ViewerDialogRegionId, string> = {
	'dialog.options':
		'The body of File > Options (the Settings dialog). Context adds `options`, `setOption`, `resetOptions`.',
	'dialog.share': 'The body of the Share dialog.',
	'dialog.broadcast': 'The body of the Broadcast dialog.',
	'dialog.print': 'The body of the Print dialog.',
};

export const VIEWER_EXPORT_FORMAT_DESCRIPTIONS: Record<ViewerExportFormatId, string> = {
	pdf: 'Export to PDF.',
	png: 'Export the current slide as PNG.',
	video: 'Export the deck as a video.',
	gif: 'Export the deck as an animated GIF.',
	json: 'Export the parsed deck as JSON.',
	copyImage: 'Copy the current slide to the clipboard as an image.',
};

export const EDITOR_SHORTCUT_DESCRIPTIONS: Record<EditorKeyActionName, string> = {
	undo: 'Undo (Ctrl/Cmd+Z).',
	redo: 'Redo (Ctrl/Cmd+Y, Ctrl/Cmd+Shift+Z).',
	copy: 'Copy the selection (Ctrl/Cmd+C).',
	cut: 'Cut the selection (Ctrl/Cmd+X).',
	paste: 'Paste (Ctrl/Cmd+V).',
	duplicate: 'Duplicate the selection (Ctrl/Cmd+D).',
	delete: 'Delete the selection (Delete, Backspace).',
	selectAll: 'Select every element on the slide (Ctrl/Cmd+A).',
	group: 'Group the selection (Ctrl/Cmd+G).',
	ungroup: 'Ungroup (Ctrl/Cmd+Shift+G).',
	nudge: 'Move the selection with the arrow keys (disable only; not remappable).',
	prevSlide: 'Previous slide (ArrowLeft with nothing selected).',
	nextSlide: 'Next slide (ArrowRight with nothing selected).',
	escape: 'Escape: leave the current mode (disable only; not remappable).',
	find: 'Find (Ctrl/Cmd+F).',
	findReplace: 'Find and replace (Ctrl/Cmd+H).',
	toggleShortcuts: 'Keyboard-shortcut reference (?, Ctrl/Cmd+/).',
	alignLeft: 'Align text left (Ctrl/Cmd+L).',
	alignCenter: 'Center text (Ctrl/Cmd+E).',
	alignRight: 'Align text right (Ctrl/Cmd+R).',
	alignJustify: 'Justify text (Ctrl/Cmd+J).',
	increaseFontSize: 'Increase font size (Ctrl/Cmd+Shift+>).',
	decreaseFontSize: 'Decrease font size (Ctrl/Cmd+Shift+<).',
	copyFormat: 'Copy formatting (Ctrl/Cmd+Shift+C).',
	pasteFormat: 'Paste formatting (Ctrl/Cmd+Shift+V).',
	newSlide: 'New slide (Ctrl/Cmd+M).',
	hyperlink: 'Insert or edit a hyperlink (Ctrl/Cmd+K).',
	clearFormatting: 'Clear character formatting (Ctrl/Cmd+Space).',
	cycleSelectionNext: 'Select the next element (Tab).',
	cycleSelectionPrev: 'Select the previous element (Shift+Tab).',
	pasteSpecial: 'Paste Special (Ctrl/Cmd+Alt+V).',
};
