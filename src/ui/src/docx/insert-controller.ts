import { applyParagraphFormat } from './paragraph-format';
import { createRuler } from './ruler';
import { syncRuler } from './ruler-sync';
import { keymap } from 'prosemirror-keymap';
import type { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { DocumentModel, DrawingWrap, PendingMediaPart } from 'ooxml-core/docx';
import { applyCharacterStyle } from './character-style-picker';
import { createLinkDialog, type LinkDialog } from './link-dialog';
import { createPictureDialog, type PictureDialog } from './picture-dialog';
import { NodeSelection } from 'prosemirror-state';
import { followLinkAt } from './link-commands';
import { insertContentControl } from 'ooxml-core/docx/ui';
import { EQUATION_GALLERY } from 'ooxml-core/docx/ui';
import { ShapeTool, shapeDrawPlugin } from './shape-draw';
import {
	setShapeFill,
	setShapeOutline,
	setShapeSize,
	setShapeWeight,
	setShapeWrap,
	type ShapeWrap,
} from './shape-commands';
import {
	insertPictures,
	newPicturePartName,
	PICTURE_TYPES,
	readBytes,
	stagePicture,
} from './picture-commands';
import { applyPictureBulletList } from './list-commands';
import {
	pictureOf,
	selectedPicturePos,
	setGraphicsFillAt,
	setPictureWrap,
} from './picture-format-commands';
import { pictureClipboardPlugin } from './picture-drop';
import { SmartArtController } from './smartart-controller';
import {
	setSelectedChartLegend,
	switchSelectedChartRowColumn,
	toggleSelectedChartElement,
} from './chart-commands';
import type { RibbonAction } from './ribbon';
import { translate, translateUiText, type EditorLocale } from './localization';
import { pickOfficeSymbol } from '../form/symbol-picker';
import { insertPlainText } from './insert-text-commands';
import { insertLocalDrawing } from './insert-local-drawing';
import { captureScreenshotFile } from '../form/display-clipping';
import { pickOfficeIcon } from '../form/icon-picker';
import { openOfficeAddIns } from '../ribbon/add-in-dialog';
import { insertTextBox, selectedTextBox, updateTextBox } from './text-box-commands';
import { focusView } from './focus-view';
import { updateFields } from './field-update';
import {
	blockPageNumbers,
	insertTableOfContents,
	removeTableOfContents,
	updateTableOfContents,
} from './toc-commands';

export interface InsertControllerHost {
	view(): EditorView | undefined;
	model(): DocumentModel;
	/** The editable content width in CSS pixels, used to fit inserted pictures. */
	contentWidth(): number;
	paper(): HTMLElement | undefined;
	toolbar(): HTMLElement | undefined;
	reportError(error: Error): void;
	/** Refreshes the ribbon and panes after a change that is not a document transaction. */
	refresh?(): void;
	/** Shared media hosts allocate immutable room-wide names and publish bytes before insertion. */
	newMediaPartName?(contentType: string): string;
	stageMedia?(partName: string, part: PendingMediaPart): void;
}

/** Picture, link, character style and hidden-text commands shared by every editor instance. */
export class InsertController {
	readonly linkDialog: LinkDialog;
	readonly pictureDialog: PictureDialog;
	readonly pictureInput: HTMLInputElement;
	/** File chooser behind Bullets > Picture bullet. */
	readonly bulletInput: HTMLInputElement;
	/** Bytes for pictures inserted since the document was loaded; passed to save. */
	readonly pendingMedia = new Map<string, PendingMediaPart>();
	/** The shape Insert > Shapes armed for the next drag on the page. */
	readonly shapeTool = new ShapeTool();
	/** Insert > SmartArt, SmartArt Design and the SmartArt text pane. */
	readonly smartArt: SmartArtController;
	private locale: EditorLocale = 'en';
	private showHidden = false;
	private showMarks = false;
	private showGridlines = false;
	private showRuler = false;

	constructor(private readonly host: InsertControllerHost) {
		this.linkDialog = createLinkDialog({
			getView: () => host.view(),
			onError: (error) => host.reportError(error),
		});
		this.pictureDialog = createPictureDialog(() => host.view());
		this.pictureInput = document.createElement('input');
		this.pictureInput.type = 'file';
		this.pictureInput.accept = [...Object.keys(PICTURE_TYPES), 'image/svg+xml'].join(',');
		this.pictureInput.hidden = true;
		this.pictureInput.className = 'dve-picture-input';
		this.pictureInput.addEventListener('change', () => void this.insertSelectedPicture());
		this.bulletInput = document.createElement('input');
		this.bulletInput.type = 'file';
		this.bulletInput.accept = Object.keys(PICTURE_TYPES).join(',');
		this.bulletInput.hidden = true;
		this.bulletInput.className = 'dve-picture-bullet-input';
		this.bulletInput.addEventListener('change', () => void this.applySelectedPictureBullet());
		this.smartArt = new SmartArtController({
			view: () => host.view(),
			contentWidth: () => host.contentWidth(),
			stage: (partName, part) => {
				this.pendingMedia.set(partName, part);
				host.stageMedia?.(partName, part);
			},
			reportError: (error) => host.reportError(error),
			refresh: () => host.refresh?.(),
		});
	}

	/** Handles the actions this controller owns; returns false for everything else. */
	handle(action: RibbonAction): boolean {
		const view = this.host.view();
		const editable = Boolean(view?.editable);
		if (this.smartArt.handle(action)) return true;
		if (action.type === 'insertPicture') {
			if (editable) this.pictureInput.click();
		} else if (action.type === 'insertScreenshot') {
			if (view?.editable) void this.insertScreenshot(view, action.clip);
		} else if (action.type === 'openSymbolGallery') {
			if (view?.editable) void this.insertSymbol(view);
		} else if (action.type === 'officeAddIns') {
			const toolbar = this.host.toolbar();
			const root = toolbar?.getRootNode();
			if (toolbar && root?.nodeType === 11) {
				openOfficeAddIns(
					root as ShadowRoot,
					action.mode,
					(label) => translateUiText(toolbar, label),
					'word',
					this.locale,
				);
			}
		} else if (action.type === 'insertIcon') {
			if (view?.editable) void this.insertIcon(view);
		} else if (action.type === 'insertLocalDrawing') {
			if (view?.editable)
				void insertLocalDrawing(view, action.kind, this.locale, this.host, this.pendingMedia);
		} else if (action.type === 'insertWordArt') {
			if (view?.editable) void this.insertWordArt(view);
		} else if (action.type === 'pictureBullet') {
			if (editable) this.bulletInput.click();
		} else if (action.type === 'insertContentControl') {
			const tr = view?.editable && insertContentControl(view.state, action.value);
			if (view && tr) view.dispatch(tr.scrollIntoView());
			focusView(view);
		} else if (action.type === 'drawShape') {
			if (editable) this.shapeTool.arm(action.preset);
		} else if (action.type === 'shape') {
			if (view) this.runShapeAction(view, action);
			focusView(view);
		} else if (action.type === 'chart' && action.key === 'legend') {
			if (view) setSelectedChartLegend(view, action.value);
			focusView(view);
		} else if (action.type === 'chart' && action.key === 'addElement') {
			if (view) toggleSelectedChartElement(view, action.value);
			focusView(view);
		} else if (action.type === 'chart' && action.key === 'switchRowColumn') {
			if (view) switchSelectedChartRowColumn(view);
			focusView(view);
		} else if (action.type === 'insertEquation') {
			const entry = EQUATION_GALLERY.find((item) => item.id === action.id);
			if (view?.editable && entry) {
				const selection = view.state.selection;
				void import('./equation-commands').then(
					({ insertEquation }) => {
						if (this.host.view() === view && selection.eq(view.state.selection)) {
							insertEquation(view, entry.latex);
							focusView(view);
						}
					},
					(cause: unknown) =>
						this.host.reportError(cause instanceof Error ? cause : new Error(String(cause))),
				);
			}
		} else if (action.type === 'link') {
			if (editable) this.linkDialog.open();
		} else if (action.type === 'formatPicture') {
			const selection = view?.state.selection;
			if (
				editable &&
				selection instanceof NodeSelection &&
				selection.node.type.name === 'image' &&
				!selection.node.attrs.payload
			)
				this.pictureDialog.open(selection.from);
		} else if (action.type === 'picture') {
			const pos = view && selectedPicturePos(view.state);
			if (!view || !editable || pos === undefined) return true;
			if (action.key === 'crop') this.pictureDialog.open(pos, 'crop');
			else if (action.key === 'graphicsFill') void this.graphicsFill(view, pos);
			else {
				setPictureWrap(view, action.value as DrawingWrap);
				focusView(view);
			}
		} else if (action.type === 'characterStyle') {
			if (view?.editable) applyCharacterStyle(view, action.value);
			focusView(view);
		} else if (action.type === 'showHidden') this.setShowHidden(!this.showHidden);
		else if (action.type === 'ruler') {
			this.showRuler = !this.showRuler;
			this.syncPaper();
		} else if (action.type === 'gridlines') {
			this.showGridlines = !this.showGridlines;
			this.syncPaper();
		} else if (action.type === 'showMarks') {
			this.showMarks = !this.showMarks;
			this.syncPaper();
		} else if (action.type === 'toc') {
			if (!view || !editable) return true;
			try {
				if (action.key === 'figures')
					insertTableOfContents(view, this.host.model(), undefined, 3, action.label ?? 'Figure');
				else if (action.key === 'updateFigures') {
					const labels = action.label ? [action.label] : ['Figure', 'Table', 'Equation'];
					// Each update rebuilds the model from the view, so re-read it between labels.
					const updated = labels.filter((label) =>
						updateTableOfContents(view, this.host.model(), undefined, label),
					);
					if (!updated.length)
						this.host.reportError(new Error('This document has no table of figures to update.'));
				} else if (action.key === 'insert')
					insertTableOfContents(view, this.host.model(), undefined, action.levels ?? 3);
				else if (action.key === 'remove') {
					if (!removeTableOfContents(view, this.host.model()))
						this.host.reportError(new Error('This document has no table of contents to remove.'));
				} else if (action.key === 'fields') {
					// Word's Update Field refreshes everything: captions and references first, then
					// the tables built from them (headings, then each figure label).
					let changed = updateFields(
						view,
						(id) => blockPageNumbers(this.host.model()).get(id),
						this.host.model(),
					);
					for (const label of [undefined, 'Figure', 'Table', 'Equation'])
						changed = updateTableOfContents(view, this.host.model(), undefined, label) || changed;
					if (!changed) this.host.reportError(new Error('No fields needed updating.'));
				} else if (!updateTableOfContents(view, this.host.model()))
					this.host.reportError(new Error('This document has no table of contents to update.'));
			} catch (error) {
				this.host.reportError(error instanceof Error ? error : new Error(String(error)));
			}
			focusView(view);
		} else return false;
		return true;
	}

	/** Ctrl+K opens the link dialog; Ctrl/Cmd+Click follows a link like Word. */
	plugins(): Plugin[] {
		return [
			shapeDrawPlugin(this.shapeTool),
			pictureClipboardPlugin(
				(view) => view === this.host.view(),
				(file) => this.insertPictureFile(file),
			),
			keymap({
				'Mod-k': (_state, _dispatch, view) => {
					if (view?.editable) this.linkDialog.open();
					return true;
				},
			}),
		];
	}

	private runShapeAction(view: EditorView, action: Extract<RibbonAction, { type: 'shape' }>): void {
		const { key, value } = action;
		if (key === 'fill') setShapeFill(view, String(value));
		else if (key === 'outline') setShapeOutline(view, String(value));
		else if (key === 'weight') setShapeWeight(view, Number(value));
		else if (key === 'width') setShapeSize(view, Number(value) * 96, undefined);
		else if (key === 'height') setShapeSize(view, undefined, Number(value) * 96);
		else if (key === 'wrap') setShapeWrap(view, String(value) as ShapeWrap);
	}

	handleClick(view: EditorView, pos: number, event: MouseEvent): boolean {
		return (event.ctrlKey || event.metaKey) && followLinkAt(view, pos);
	}

	/** Media bytes for rendering: pictures inserted in this session first, then the loaded package. */
	media(partName: string, loaded?: ReadonlyMap<string, Uint8Array>): Uint8Array | undefined {
		return this.pendingMedia.get(partName)?.bytes ?? loaded?.get(partName);
	}

	reset(): void {
		this.shapeTool.disarm();
		this.pendingMedia.clear();
		this.smartArt.reset();
		if (this.linkDialog.isOpen) this.linkDialog.close();
		if (this.pictureDialog.isOpen) this.pictureDialog.close();
	}

	setLocale(locale: EditorLocale): void {
		this.locale = locale;
		this.linkDialog.setLocale(locale);
		this.pictureDialog.setLocale(locale);
		this.smartArt.setLocale(locale);
		this.pictureInput.setAttribute('aria-label', translate(locale, 'Insert picture'));
		this.bulletInput.setAttribute('aria-label', translate(locale, 'Picture bullet...'));
	}

	/** Re-applies the hidden-text display state after the paper is rebuilt. */
	syncPaper(): void {
		const paper = this.host.paper();
		if (paper) {
			paper.toggleAttribute('data-show-hidden', this.showHidden);
			paper.toggleAttribute('data-show-marks', this.showMarks);
			paper.toggleAttribute('data-show-gridlines', this.showGridlines);
			this.syncRulerElement(paper);
		}
		this.host
			.toolbar()
			?.querySelector(
				'[aria-label="Show paragraph marks"], [data-localearialabel="Show paragraph marks"]',
			)
			?.setAttribute('aria-pressed', String(this.showMarks));
		this.host
			.toolbar()
			?.querySelector('[aria-label="Gridlines"], [data-localearialabel="Gridlines"]')
			?.setAttribute('aria-pressed', String(this.showGridlines));
		this.host
			.toolbar()
			?.querySelector('[aria-label="Show hidden text"], [data-localearialabel="Show hidden text"]')
			?.setAttribute('aria-pressed', String(this.showHidden));
	}

	/** The ruler sits above the page in the canvas while View > Ruler is on. */
	private syncRulerElement(paper: HTMLElement): void {
		const canvas = paper.parentElement;
		const existing = canvas?.querySelector('.dve-ruler');
		if (this.showRuler && canvas && !existing)
			canvas.insertBefore(
				createRuler((change) => this.applyRulerChange(change)),
				paper,
			);
		else if (!this.showRuler) existing?.remove();
		const frame = this.host.toolbar()?.parentElement;
		const button = frame?.querySelector('[aria-label="Ruler"], [data-localearialabel="Ruler"]');
		button?.setAttribute('aria-pressed', String(this.showRuler));
		const view = this.host.view();
		const toolbar = this.host.toolbar();
		if (this.showRuler && view && toolbar) syncRuler(toolbar, view, this.host.model());
	}

	private applyRulerChange(
		change: Parameters<NonNullable<Parameters<typeof createRuler>[0]>>[0],
	): void {
		const view = this.host.view();
		const toolbar = this.host.toolbar();
		if (view && Object.keys(change).length) applyParagraphFormat(view, change);
		if (view && toolbar) syncRuler(toolbar, view, this.host.model());
		focusView(view);
	}

	private setShowHidden(show: boolean): void {
		this.showHidden = show;
		this.syncPaper();
	}

	private async insertSymbol(view: EditorView): Promise<void> {
		const host = this.host.toolbar();
		if (!host) return;
		const symbol = await pickOfficeSymbol(host, {
			heading: translate(this.locale, 'Symbol'),
			close: translate(this.locale, 'Close'),
			font: translate(this.locale, 'Font'),
		});
		if (symbol && this.host.view() === view && view.editable) {
			insertPlainText(view, symbol);
			focusView(view);
		}
	}

	private async insertIcon(view: EditorView): Promise<void> {
		const host = this.host.toolbar();
		if (!host) return;
		const t = (label: import('./localization').LocalizationKey) => translate(this.locale, label);
		const icons = await pickOfficeIcon(host, {
			heading: t('Icons'),
			search: t('Search icons'),
			color: t('Icon color'),
			empty: t('No matching icons.'),
			cancel: t('Cancel'),
			insert: t('Insert'),
			close: t('Close'),
		});
		if (this.host.view() === view && view.editable) {
			await this.insertPictureFiles(
				icons.map((icon) => new File([icon.svg], `${icon.name}.svg`, { type: 'image/svg+xml' })),
			);
		}
	}

	private async graphicsFill(view: EditorView, pos: number): Promise<void> {
		const host = this.host.toolbar();
		const node = view.state.doc.nodeAt(pos);
		if (!host || !node?.attrs.svgPartName) return;
		const t = (label: import('./localization').LocalizationKey) => translate(this.locale, label);
		const { pickGraphicsFill } = await import('../form/graphics-fill');
		if (this.host.view() !== view || !view.editable) return;
		const patch = await pickGraphicsFill(host, pictureOf(node).pictureFill?.effects, {
			heading: t('Graphics Fill'),
			color: t('Color'),
			reset: t('No Recolor'),
			cancel: t('Cancel'),
			close: t('Close'),
			apply: t('OK'),
		});
		if (!patch || this.host.view() !== view || !view.editable) return;
		if (selectedPicturePos(view.state) === pos && view.state.doc.nodeAt(pos) === node)
			setGraphicsFillAt(view, pos, patch);
	}

	private async insertWordArt(view: EditorView): Promise<void> {
		const host = this.host.toolbar();
		if (!host) return;
		const t = (label: import('./localization').LocalizationKey) => translate(this.locale, label);
		const box = selectedTextBox(view.state);
		const current = box?.settings.wordArt ? box : undefined;
		const currentNode = current ? view.state.doc.nodeAt(current.pos) : undefined;
		const { pickWordArt } = await import('../form/word-art-picker');
		if (this.host.view() !== view || !view.editable) return;
		const selected = await pickWordArt(
			host,
			{
				heading: t('WordArt'),
				text: t('WordArt text'),
				placeholder: t('Your text here'),
				style: (index) => `${t('WordArt style')} ${index}`,
				transform: t('Transform'),
				preview: t('Preview'),
				locale: this.locale,
				confirm: current ? t('OK') : t('Insert'),
				cancel: t('Cancel'),
				close: t('Close'),
			},
			current?.settings.wordArt
				? { text: current.settings.lines.join(' '), style: current.settings.wordArt }
				: undefined,
		);
		if (!selected || this.host.view() !== view || !view.editable) return;
		if (current) {
			if (view.state.doc.nodeAt(current.pos) === currentNode) {
				updateTextBox(view, current.pos, {
					...current.settings,
					lines: [selected.text],
					wordArt: selected.style,
				});
			}
			return;
		}
		insertTextBox(view, {
			lines: [selected.text],
			widthPx: Math.min(550, Math.max(180, selected.text.length * 22 + 26)),
			heightPx: selected.style.warp ? 115 : 70,
			border: false,
			wordArt: selected.style,
		});
	}

	private async insertScreenshot(view: EditorView, clip: boolean): Promise<void> {
		const parent = this.host.toolbar() ?? this.host.paper();
		if (!parent) return;
		const selection = view.state.selection;
		const t = (key: Parameters<typeof translate>[1]) => translate(this.locale, key);
		try {
			const file = await captureScreenshotFile(
				parent,
				t('Screenshot'),
				clip
					? {
							heading: t('Screen Clipping'),
							close: t('Close'),
							cancel: t('Cancel'),
							insert: t('Insert'),
							x: 'X',
							y: 'Y',
							width: t('Width'),
							height: t('Height'),
						}
					: undefined,
			);
			if (file && this.host.view() === view && view.editable && selection.eq(view.state.selection))
				await this.insertPictureFile(file);
		} catch (cause) {
			this.host.reportError(cause instanceof Error ? cause : new Error(String(cause)));
		}
	}

	private async insertSelectedPicture(): Promise<void> {
		const file = this.pictureInput.files?.[0];
		this.pictureInput.value = '';
		if (file) await this.insertPictureFile(file);
	}

	private async applySelectedPictureBullet(): Promise<void> {
		const file = this.bulletInput.files?.[0];
		this.bulletInput.value = '';
		if (file) await this.applyPictureBulletFile(file);
	}

	/** Stages `file` as a new picture bullet and starts a list that uses it. */
	async applyPictureBulletFile(file: File): Promise<boolean> {
		const view = this.host.view();
		if (!view?.editable) return false;
		try {
			const contentType = file.type.toLowerCase();
			if (!PICTURE_TYPES[contentType])
				throw new Error('Choose a PNG, JPEG, GIF or BMP picture for a picture bullet.');
			const bytes = await readBytes(file);
			if (this.host.view() !== view || !view.editable) return false;
			const partName = (this.host.newMediaPartName ?? newPicturePartName)(contentType);
			const part = { bytes, contentType };
			this.host.stageMedia?.(partName, part);
			this.pendingMedia.set(partName, part);
			applyPictureBulletList(view, this.host.model(), { partName, contentType });
			this.host.refresh?.();
			focusView(view);
			return true;
		} catch (cause) {
			this.host.reportError(cause instanceof Error ? cause : new Error(String(cause)));
			return false;
		}
	}

	/** Stages `file` and inserts it at the selection; false when it could not be inserted. */
	insertPictureFile(file: File): Promise<boolean> {
		return this.insertPictureFiles([file]);
	}

	private async insertPictureFiles(files: readonly File[]): Promise<boolean> {
		const view = this.host.view();
		if (!view?.editable || !files.length) return false;
		try {
			const staged = await Promise.all(
				files.map((file) =>
					stagePicture(file, this.host.contentWidth(), this.host.newMediaPartName),
				),
			);
			if (this.host.view() !== view || !view.editable) return false;
			for (const { image, media, svg } of staged) {
				this.host.stageMedia?.(image.partName, media);
				this.pendingMedia.set(image.partName, media);
				if (svg && image.svgPartName) {
					this.host.stageMedia?.(image.svgPartName, svg);
					this.pendingMedia.set(image.svgPartName, svg);
				}
			}
			const inserted = insertPictures(
				view,
				staged.map(({ image }) => image),
			);
			focusView(view);
			return inserted;
		} catch (cause) {
			this.host.reportError(cause instanceof Error ? cause : new Error(String(cause)));
			return false;
		}
	}
}
