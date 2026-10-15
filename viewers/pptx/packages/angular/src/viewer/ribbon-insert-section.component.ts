/**
 * ribbon-insert-section.component.ts: the Insert ribbon tab, split out of
 * {@link RibbonComponent}. A thin adapter for the shared `pptx-ui-ribbon-insert`:
 * the shared element owns groups, icons, labels, pickers and pressed/disabled state;
 * this component supplies viewer state and routes typed intents to the native
 * handlers.
 *
 * Document mutation stays native: inserts go straight through the shared
 * {@link EditorStateService}; the file dialog / FileReader / image-probe plumbing
 * lives in `ribbon-insert-file-picker.ts`; Action buttons, Fields and the Date/Time
 * modal live in {@link RibbonInsertFieldsComponent}; SmartArt, Equation and Hyperlink
 * open the viewer's own dialogs through outputs, and Header & Footer through
 * {@link ViewerDialogsService}.
 *
 * The chart-type and shape-type selections are owned by the parent ribbon (so they
 * survive a tab switch) and passed in via `newChartType` / `newShapeType`; changes
 * emit `chartTypeChange` / `shapeTypeChange`.
 */
import {
	ChangeDetectionStrategy,
	Component,
	computed,
	CUSTOM_ELEMENTS_SCHEMA,
	inject,
	input,
	output,
	viewChild,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import {
	classifyMediaType,
	DEFAULT_INSERT_CHART_KIND,
	resolveInsertSlideLayout,
} from 'ooxml-ui/pptx';
import type {
	FreeformToolKind,
	InsertChartKind,
	RibbonInsertRequestEvent,
	ShapePresetType,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import {
	newChartElement,
	newPresetShapeElement,
	newTableElement,
	newTextElement,
} from './editor-insert';
import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { OutlineAuthoringService } from './outline-authoring.service';
import { RibbonInsertFieldsComponent } from './ribbon-insert-fields.component';
import { imageDimensions, pickFile, readAsDataUrl } from './ribbon-insert-file-picker';
import { visibleFreeformTools } from './ribbon-insert-freeform';
import { layoutOptionsFrom } from './ribbon-layout-options';
import { injectResolvedCustomization } from './viewer-customization.service';
import { ViewerDialogsService } from './viewer-dialogs.service';

@Component({
	selector: 'pptx-ribbon-insert-section',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	imports: [RibbonInsertFieldsComponent],
	template: `
		@if (phone()) {
			<pptx-ui-mobile-insert [state]="view()" (insert-request)="request($event)" />
		} @else {
			<pptx-ui-ribbon-insert [state]="view()" (insert-request)="request($event)" />
		}
		<pptx-ribbon-insert-fields [slideIndex]="slideIndex()" />
	`,
})
export class RibbonInsertSectionComponent {
	private readonly editor = inject(EditorStateService);
	private readonly translate = inject(TranslateService);
	private readonly dialogs = inject(ViewerDialogsService);
	private readonly outline = inject(OutlineAuthoringService, { optional: true });
	private readonly customization = injectResolvedCustomization();
	private readonly fields = viewChild(RibbonInsertFieldsComponent);
	private readonly loader = inject(LoadContentService);
	/** Layouts New Slide offers: one entry per layout of the deck's masters. */
	private readonly layouts = computed(() => layoutOptionsFrom(this.loader.slideMasters()));

	readonly slideIndex = input<number>(0);
	/** Whether the comments pane is open: Comment must not toggle it shut. */
	readonly commentsOpen = input<boolean>(false);
	readonly canEdit = input<boolean>(false);
	/** The insert-chart dropdown entry ('column' is vertical, 'bar' horizontal). */
	readonly newChartType = input<InsertChartKind>(DEFAULT_INSERT_CHART_KIND);
	readonly newShapeType = input<ShapePresetType>('rect');
	/** Draw the phone Insert sheet (`pptx-ui-mobile-insert`) instead of the ribbon group row. */
	readonly phone = input<boolean>(false);
	/** Phone only: a request was handled, so the sheet can close. */
	readonly handled = output<void>();

	readonly openSmartArtDialog = output<void>();
	readonly openEquationDialog = output<void>();
	/** "Hyperlink"; the host opens the hyperlink edit dialog for the selection. */
	readonly openHyperlink = output<void>();
	/** "Comment"; the host opens the comments pane, as Review > New Comment does. */
	readonly toggleComments = output<void>();
	readonly chartTypeChange = output<InsertChartKind>();
	readonly shapeTypeChange = output<ShapePresetType>();

	protected view() {
		return {
			editable: this.canEdit(),
			hasSelection: this.editor.hasSelection(),
			shapeType: this.newShapeType(),
			chartKind: this.newChartType(),
			activeFreeformTool: this.outline?.activeFreeformTool() ?? null,
			freeformTools: visibleFreeformTools(this.outline, this.customization()),
			// Like Home, New Slide works without layouts (it adds a blank slide).
			newSlideAvailable: true,
			commentAvailable: true,
			layouts: this.layouts(),
			...this.currentLayout(),
			translate: (key: string) => this.translate.instant(key),
		};
	}

	/** The active slide's layout, marked in the New Slide gallery. */
	private currentLayout(): { currentLayoutPath?: string } {
		const path = this.editor.slides()[this.slideIndex()]?.layoutPath;
		return path ? { currentLayoutPath: path } : {};
	}

	protected request(event: Event): void {
		const intent = (event as RibbonInsertRequestEvent).detail;
		switch (intent.kind) {
			case 'command': {
				const commands = {
					textBox: () => this.editor.addElement(this.slideIndex(), newTextElement()),
					table: () => this.editor.addElement(this.slideIndex(), newTableElement()),
					image: () => this.insertImage(),
					media: () => this.insertMedia(),
					smartArt: () => this.openSmartArtDialog.emit(),
					equation: () => this.openEquationDialog.emit(),
					link: () => this.openHyperlink.emit(),
					headerFooter: () => this.dialogs.showHeaderFooter.set(true),
					// Opening an already open pane would toggle it shut.
					comment: () => (this.commentsOpen() ? undefined : this.toggleComments.emit()),
				};
				commands[intent.value]();
				break;
			}
			case 'table':
				this.editor.addElement(this.slideIndex(), newTableElement(intent.rows, intent.columns));
				break;
			case 'newSlide': {
				const layout = resolveInsertSlideLayout(
					this.layouts(),
					intent.value,
					this.currentLayout().currentLayoutPath,
				);
				this.editor.addSlide(this.slideIndex(), layout?.path);
				break;
			}
			case 'shapeType':
				this.shapeTypeChange.emit(intent.value as ShapePresetType);
				break;
			case 'shape':
				this.editor.addElement(
					this.slideIndex(),
					newPresetShapeElement(intent.value as ShapePresetType),
				);
				break;
			case 'chartType':
				this.chartTypeChange.emit(intent.value as InsertChartKind);
				break;
			case 'chart':
				this.editor.addElement(this.slideIndex(), newChartElement(intent.value as InsertChartKind));
				break;
			case 'freeform':
				this.outline?.armFreeformTool(intent.value as FreeformToolKind | null);
				break;
			case 'actionButton':
				this.fields()?.addActionButton(intent.value);
				break;
			case 'field':
				if (intent.value === 'datetime') {
					this.fields()?.openDatePicker();
				} else {
					this.fields()?.insertField(intent.value);
				}
		}
		if (this.phone()) {
			this.handled.emit();
		}
	}

	/** Pick an image file and add it as an inline image element (data-URL backed). */
	private insertImage(): void {
		pickFile('image/*', (file) => void this.addImageFile(file));
	}

	/** Pick an audio/video file and add it as a media element (data-URL backed). */
	private insertMedia(): void {
		pickFile('video/*,audio/*', (file) => void this.addMediaFile(file));
	}

	private async addImageFile(file: File): Promise<void> {
		const dataUrl = await readAsDataUrl(file);
		if (!dataUrl) {
			return;
		}
		const dims = await imageDimensions(dataUrl);
		const maxW = 400;
		const scale = dims.width > maxW ? maxW / dims.width : 1;
		const element: PptxElement = {
			type: 'image',
			id: '',
			name: file.name || this.translate.instant('pptx.elementType.image'),
			x: 100,
			y: 100,
			width: Math.round(dims.width * scale),
			height: Math.round(dims.height * scale),
			imageData: dataUrl,
		} as PptxElement;
		this.editor.addElement(this.slideIndex(), element);
	}

	private async addMediaFile(file: File): Promise<void> {
		const dataUrl = await readAsDataUrl(file);
		if (!dataUrl) {
			return;
		}
		const mediaType = classifyMediaType(file.type);
		if (!mediaType) {
			return;
		}
		const isAudio = mediaType === 'audio';
		const element: PptxElement = {
			type: 'media',
			id: '',
			name: file.name || this.translate.instant('pptx.elementType.media'),
			x: 100,
			y: 100,
			width: isAudio ? 280 : 480,
			height: isAudio ? 64 : 270,
			mediaType,
			mediaData: dataUrl,
			mediaMimeType: file.type,
		} as PptxElement;
		this.editor.addElement(this.slideIndex(), element);
	}
}
