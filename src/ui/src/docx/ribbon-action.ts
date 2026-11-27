import type { DocumentView, OutlineAction } from './document-views';
import type { TableCommand } from './table-commands';
import type { TableTabAction } from './table-ribbon';
import type { MultilingualAction } from './multilingual-ribbon';
import type { ReviewDisplayMode } from './review-display';
import type { DiagramFormatChange } from 'ooxml-core/diagram';

/** Every command the ribbon can raise, dispatched as a `ribbon-action` event. */
export type RibbonAction =
	| { type: 'documentDesign'; menu: import('./design-ribbon-controls').DesignMenu }
	| { type: 'mailings'; key: import('./mailings-ribbon').MailingsCommand }
	| { type: 'references'; key: import('./references-ribbon').ReferencesCommand }
	| { type: 'smartArtFormat'; change: DiagramFormatChange }
	| { type: 'help' }
	| {
			type: 'format';
			key: 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript';
	  }
	| { type: 'history'; key: 'undo' | 'redo' }
	| { type: 'align'; value: 'left' | 'center' | 'right' | 'justify' }
	| { type: 'font'; key: 'family' | 'size' | 'color' | 'highlight'; value: string }
	| { type: 'clear' }
	| { type: 'table'; rows?: number; columns?: number }
	| { type: 'tablePicker' }
	| { type: 'clipboard'; key: 'cut' | 'copy' | 'paste'; plainText?: boolean }
	| { type: 'underline'; value: import('./font-format').UnderlineKind }
	| { type: 'fontStep'; direction: 'grow' | 'shrink' }
	| { type: 'insertBreak'; kind: 'page' | 'column' }
	| { type: 'tableEdit'; key: TableCommand }
	| TableTabAction
	| {
			type: 'page';
			key:
				| 'margin'
				| 'orientation'
				| 'columns'
				| 'numberFormat'
				| 'numberStart'
				| 'titlePage'
				| 'verticalAlign'
				| 'lineNumbers'
				| 'size';
			value: string;
	  }
	| { type: 'sectionBreak'; kind: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' }
	| { type: 'evenOddHeaders' }
	| { type: 'pageColor'; value: string }
	| { type: 'hyphenation'; value: string }
	| { type: 'suppressLineNumbers' }
	| {
			type: 'paragraph';
			key: 'indent' | 'spacingBefore' | 'spacingAfter' | 'lineSpacing';
			value: string;
	  }
	| { type: 'zoom'; value: number }
	| {
			type: 'list';
			key:
				| 'bullet'
				| 'number'
				| 'multilevel'
				| 'outline'
				| 'headings'
				| 'article'
				| 'roman'
				| 'increaseLevel'
				| 'decreaseLevel'
				| 'remove';
	  }
	| { type: 'view'; value: DocumentView }
	| OutlineAction
	| { type: 'print' }
	| MultilingualAction
	| { type: 'search'; focus?: 'find' | 'replace' }
	| { type: 'changeCase'; value: 'sentence' | 'lower' | 'upper' | 'title' | 'toggle' }
	| { type: 'formatPainter' }
	| { type: 'ruler' }
	| { type: 'coverPage' }
	| { type: 'sort'; order: 'ascending' | 'descending' }
	| { type: 'goTo' }
	| { type: 'selectObjects' }
	| { type: 'addText'; level: 0 | 1 | 2 | 3 }
	| { type: 'blankPage' }
	| { type: 'readAloud' }
	| { type: 'gridlines' }
	| { type: 'navigation' }
	| { type: 'shading'; value: string }
	| {
			type: 'borders';
			preset:
				| 'none'
				| 'bottom'
				| 'top'
				| 'left'
				| 'right'
				| 'all'
				| 'outside'
				| 'insideH'
				| 'insideV'
				| 'horizontal';
	  }
	| {
			type: 'pageNumber';
			position: 'top' | 'bottom';
			align: 'left' | 'center' | 'right';
			style?: 'plain' | 'pageOfTotal';
	  }
	| {
			type: 'headerFooter';
			kind: 'header' | 'footer';
			slot?: 'default' | 'first' | 'even';
			design?: import('ooxml-core/docx').HeaderFooterDesign;
	  }
	| { type: 'headerFooterRemove'; kind: 'header' | 'footer' }
	| { type: 'headerFooterEdit'; kind: 'header' | 'footer' }
	| { type: 'quickPart'; command: QuickPartCommand }
	| { type: 'insertParts'; kind: InsertPartsKind }
	| { type: 'headerFooterLink' }
	| { type: 'headerFooterClose' }
	| { type: 'headerFooterNavigate'; target: 'header' | 'footer' | 'previous' | 'next' }
	| { type: 'headerFooterPosition'; kind: 'header' | 'footer'; inches: number }
	| {
			type: 'formatDialog';
			kind: import('./format-dialogs').FormatDialogKind;
	  }
	| { type: 'indent'; side: 'left' | 'right'; inches: number }
	| { type: 'zoomFit'; mode: 'actual' | 'width' | 'page' | 'pages' }
	| { type: 'wordCount' }
	| { type: 'spelling' }
	| { type: 'insertSymbol'; value: string }
	| { type: 'openSymbolGallery' }
	| { type: 'dropCap'; value: 'none' | 'drop' | 'margin' }
	| { type: 'insertDateTime'; value: 'long' | 'short' | 'time' | 'datetime' }
	| { type: 'showMarks' }
	| { type: 'selectAll' }
	| {
			type: 'review';
			key:
				| 'trackChanges'
				| 'acceptOne'
				| 'rejectOne'
				| 'acceptAll'
				| 'rejectAll'
				| 'previous'
				| 'next';
	  }
	| { type: 'reviewDisplay'; value: ReviewDisplayMode }
	| { type: 'comments'; key: 'toggle' | 'add' | 'delete' | 'previous' | 'next' }
	| { type: 'insertPicture' }
	| { type: 'insertScreenshot'; clip: boolean }
	| { type: 'insertIcon' }
	| { type: 'insertLocalDrawing'; kind: 'model3d' | 'object' }
	| { type: 'officeAddIns'; mode: 'getAddIns' | 'myAddIns' }
	| { type: 'insertWordArt' }
	/** Bullets > Picture bullet...: choose an image file and start a list that draws it. */
	| { type: 'pictureBullet' }
	| { type: 'shapePicker' }
	| { type: 'insertSmartArt' }
	| {
			type: 'smartArt';
			key:
				| 'addAfter'
				| 'addBefore'
				| 'addAbove'
				| 'addBelow'
				| 'promote'
				| 'demote'
				| 'moveUp'
				| 'moveDown'
				| 'textPane'
				| 'changeLayout';
	  }
	| { type: 'drawShape'; preset: string }
	/** The Draw tab: tools, the pen gallery, Draw with Touch, Ink to Shape and Math, Ink Replay. */
	| { type: 'draw'; key: 'tool'; value: 'select' | 'lasso' | 'eraser' }
	| { type: 'draw'; key: 'pen' | 'removePen'; index: number }
	| { type: 'draw'; key: 'addPen'; value: 'pen' | 'pencil' | 'highlighter' }
	| { type: 'draw'; key: 'penColor'; value: string }
	| { type: 'draw'; key: 'thickness'; value: number }
	| { type: 'draw'; key: 'touch' | 'ruler' | 'inkToShape' | 'inkToMath' | 'replay' }
	| {
			type: 'shape';
			key: 'fill' | 'outline' | 'weight' | 'width' | 'height' | 'wrap';
			value: string | number;
	  }
	| { type: 'chart'; key: 'legend'; value: string }
	| { type: 'chart'; key: 'addElement'; value: string }
	| { type: 'chart'; key: 'switchRowColumn' }
	/** Format Selection: the docked chart Format pane, on `target` (`chartArea` by default). */
	| { type: 'chart'; key: 'format'; target?: string }
	| { type: 'insertEquation'; id: import('ooxml-core/docx/ui').GalleryEquation['id'] }
	/** The Equation tab and Insert > Equation > Insert New Equation (Alt+=). */
	| { type: 'equation'; key: import('./equation-tab').EquationCommand; value?: string }
	| {
			type: 'insertContentControl';
			value: 'richText' | 'text' | 'checkbox' | 'dropDownList' | 'comboBox' | 'date';
	  }
	| { type: 'formatPicture' }
	/** Picture Format: Wrap Text (a `DrawingWrap`) or Crop (the Format Picture crop fields). */
	| { type: 'arrange'; key: string; value: string }
	| { type: 'picture'; key: 'wrap'; value: string }
	| { type: 'picture'; key: 'crop' }
	| { type: 'picture'; key: 'graphicsFill' }
	| { type: 'link' }
	| { type: 'characterStyle'; value: string }
	| { type: 'showHidden' }
	| { type: 'thumbnails' }
	| { type: 'insertNote'; kind: 'footnote' | 'endnote' }
	/** Convert Notes, or Convert to Endnote/Footnote on one note. */
	| { type: 'convertNotes'; selection: import('ooxml-core/docx').NoteSelection }
	| {
			type: 'toc';
			key: 'insert' | 'update' | 'remove' | 'fields' | 'figures' | 'updateFigures';
			levels?: number;
			label?: string;
	  };

export type InsertPartsKind =
	| 'signatureSetup'
	| 'saveQuickPart'
	| 'buildingBlocks'
	| 'fieldDialog'
	| 'onlineVideo';

/** Insert > Quick Parts: a document property or a gallery entry (the dialogs are `formatDialog`s). */
export type QuickPartCommand =
	| { kind: 'property'; id: import('ooxml-core/docx').DocumentPropertyId }
	| { kind: 'entry'; key: string };
