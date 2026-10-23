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
import type { DocumentStorage } from 'ooxml-core/embedding';
import type {
	DocxCollaborationOptions,
	DocxCollaborationState,
	DocxEditorElement,
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
import { angularOutputs, installHandle } from 'ooxml-ui/bindings';
import {
	EDITOR_CONTRACT,
	mountEditor,
	pickEditorProps,
	type EditorBinding,
	type EditorEventOptions,
	type EditorHandle,
} from './index';

/** The handle members, defined on the prototype from `EDITOR_CONTRACT.handle`. */
// oxlint-disable-next-line typescript/no-unsafe-declaration-merging -- members come from installHandle.
export interface WordEditorComponent extends Omit<EditorHandle, 'element'> {
	/** The `<docx-editor>` element; `undefined` before the view is initialised. */
	readonly element: DocxEditorElement | undefined;
}
// oxlint-disable-next-line typescript/no-unsafe-declaration-merging -- see the interface above.
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
	@Input() storage?: DocumentStorage | null;
	@Input() customization?: WordCustomization | null;
	@Output() documentChange = new EventEmitter<DocumentModel>();
	@Output() documentError = new EventEmitter<Error>();
	@Output() documentWarning = new EventEmitter<string>();
	@Output() pageChange = new EventEmitter<PageChangeDetail>();
	@Output() dirtyChange = new EventEmitter<boolean>();
	/** Cancelable: call `preventDefault()` on the emitted event to take a File command over. */
	@Output() fileCommand = new EventEmitter<CustomEvent<FileCommandDetail>>();
	/** Pairs with `readOnly` for `[(readOnly)]`: the host's own switch, never a role lock. */
	@Output() readOnlyChange = new EventEmitter<boolean>();
	@Output() ribbonAction = new EventEmitter<RibbonAction>();
	@Output() ribbonCustomize = new EventEmitter<readonly string[]>();
	@Output() ribbonAddIn = new EventEmitter<RibbonAddInCommandDetail>();
	@Output() collaborationChange = new EventEmitter<DocxCollaborationState>();
	private host = inject<ElementRef<HTMLElement>>(ElementRef);
	private binding?: EditorBinding;
	static {
		installHandle(this.prototype, EDITOR_CONTRACT.handle, (self) => self.binding, {
			lenientGetters: true,
		});
	}
	private options() {
		return {
			...pickEditorProps(this),
			...angularOutputs<EditorEventOptions>(EDITOR_CONTRACT, this),
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
}
