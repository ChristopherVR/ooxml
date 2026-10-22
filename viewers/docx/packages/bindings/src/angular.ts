import {
	Component,
	ElementRef,
	EventEmitter,
	Input,
	Output,
	inject,
	type AfterViewInit,
	type OnChanges,
	type OnDestroy,
} from '@angular/core';
import type { DocumentModel } from 'docx-core';
import type {
	DocxCollaborationOptions,
	DocxCollaborationState,
	EditorTheme,
	EditorThemeMode,
	FileCommandDetail,
	PageChangeDetail,
	RibbonAction,
	RibbonActionInput,
	RibbonAddInCommandDetail,
	RibbonAddInTab,
	WordCustomization,
} from 'docx-web-component';
import { eventOptions, mountEditor, pickEditorProps, type EditorBinding } from './index';
@Component({ selector: 'word-editor', standalone: true, template: '' })
export class WordEditorComponent implements AfterViewInit, OnChanges, OnDestroy {
	@Input() documentModel?: DocumentModel;
	@Input() readOnly = false;
	@Input() fileName?: string;
	@Input() reviewAuthor?: string;
	@Input() themeColors?: Partial<EditorTheme>;
	@Input() locale = 'en';
	@Input() theme: EditorThemeMode = 'auto';
	@Input() showThumbnails = false;
	@Input() showToolbar = true;
	@Input() hiddenActions: readonly RibbonActionInput[] = [];
	@Input() ribbonAddIns?: readonly RibbonAddInTab[];
	@Input() collaboration?: DocxCollaborationOptions | null;
	@Input() customization?: WordCustomization | null;
	@Output() documentChange = new EventEmitter<DocumentModel>();
	@Output() documentError = new EventEmitter<Error>();
	@Output() documentWarning = new EventEmitter<string>();
	@Output() pageChange = new EventEmitter<PageChangeDetail>();
	@Output() dirtyChange = new EventEmitter<boolean>();
	/** Cancelable: call `preventDefault()` on the emitted event to take a File command over. */
	@Output() fileCommand = new EventEmitter<CustomEvent<FileCommandDetail>>();
	@Output() readOnlyChange = new EventEmitter<boolean>();
	@Output() ribbonAction = new EventEmitter<RibbonAction>();
	@Output() ribbonCustomize = new EventEmitter<readonly string[]>();
	@Output() ribbonAddIn = new EventEmitter<RibbonAddInCommandDetail>();
	@Output() collaborationChange = new EventEmitter<DocxCollaborationState>();
	private host = inject<ElementRef<HTMLElement>>(ElementRef);
	private binding?: EditorBinding;
	private options() {
		return {
			...pickEditorProps(this),
			...eventOptions({
				'document-change': (model) => this.documentChange.emit(model),
				'document-error': (error) => this.documentError.emit(error),
				'document-warning': (message) => this.documentWarning.emit(message),
				'page-change': (detail) => this.pageChange.emit(detail),
				'dirty-change': (dirty) => this.dirtyChange.emit(dirty),
				'file-command': (event) => this.fileCommand.emit(event),
				'readonly-change': (_readOnly, detail) => {
					// Two-way binding mirrors the host's switch only, never a role lock.
					if (detail.hostReadOnly !== this.readOnly) this.readOnlyChange.emit(detail.hostReadOnly);
				},
				'ribbon-action': (detail) => this.ribbonAction.emit(detail),
				'ribbon-customize': (ids) => this.ribbonCustomize.emit(ids),
				'office-ribbon-add-in': (detail) => this.ribbonAddIn.emit(detail),
				'collaboration-change': (state) => this.collaborationChange.emit(state),
			}),
		};
	}
	ngAfterViewInit() {
		this.binding = mountEditor(this.host.nativeElement, this.options());
	}
	ngOnChanges() {
		this.binding?.update(this.options());
	}
	ngOnDestroy() {
		this.binding?.destroy();
	}
	get element() {
		return this.binding?.element;
	}
	async load(input: Uint8Array | ArrayBuffer) {
		if (!this.binding) throw new Error('Editor is not mounted');
		await this.binding.load(input);
	}
	async save() {
		if (!this.binding) throw new Error('Editor is not mounted');
		return this.binding.save();
	}
	async download(fileName?: string) {
		if (!this.binding) throw new Error('Editor is not mounted');
		await this.binding.download(fileName);
	}
	markClean() {
		this.binding?.markClean();
	}
	get dirty() {
		return this.binding?.dirty ?? false;
	}
}
