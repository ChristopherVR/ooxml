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
import type { Workbook } from '@christophervr/xlsx-core';
import type { DocumentStorage } from 'ooxml-core/embedding';
import type {
	EditorThemeMode,
	FileCommandDetail,
	RibbonAddInCommandDetail,
	RibbonAddInTab,
	SelectionChangeDetail,
	XlsxCollaborationOptions,
	XlsxCollaborationState,
	XlsxEditorElement,
	XlsxThemeColors,
	ExcelCustomization,
} from 'xlsx-web-component';
import { angularOutputs, installHandle } from 'ooxml-ui/bindings';
import {
	EDITOR_CALLBACKS,
	EDITOR_CONTRACT,
	mountEditor,
	pickEditorProps,
	type EditorBinding,
	type EditorEventOptions,
	type EditorHandle,
} from './index';

/** The handle members, defined on the prototype from `EDITOR_CONTRACT.handle`. */
// oxlint-disable-next-line typescript/no-unsafe-declaration-merging -- members come from installHandle.
export interface SpreadsheetEditorComponent extends Omit<EditorHandle, 'element'> {
	/** The `<xlsx-editor>` element; `undefined` before the view is initialised. */
	readonly element: XlsxEditorElement | undefined;
}
// oxlint-disable-next-line typescript/no-unsafe-declaration-merging -- see the interface above.
@Component({ selector: 'spreadsheet-editor', standalone: true, template: '' })
export class SpreadsheetEditorComponent implements AfterViewInit, OnChanges, OnDestroy {
	@Input() workbook?: Workbook;
	@Input() bytes?: Uint8Array | ArrayBuffer;
	@Input() src?: string;
	@Input() fileName?: string;
	@Input() readOnly = false;
	@Input() locale = 'en';
	@Input() theme: EditorThemeMode = 'auto';
	@Input() authorName = 'Author';
	@Input() showToolbar = true;
	@Input() showFormulaBar = true;
	@Input() hiddenActions: readonly string[] = [];
	@Input() ribbonAddIns?: readonly RibbonAddInTab[];
	@Input() themeColors?: XlsxThemeColors;
	@Input() collaboration?: XlsxCollaborationOptions | null;
	@Input() storage?: DocumentStorage | null;
	@Input() customization?: ExcelCustomization | null;
	@Output() workbookChange = new EventEmitter<Workbook>();
	@Output() workbookError = new EventEmitter<Error>();
	@Output() selectionChange = new EventEmitter<SelectionChangeDetail>();
	@Output() dirtyChange = new EventEmitter<boolean>();
	/**
	 * Pairs with `readOnly` for `[(readOnly)]` two-way binding. It carries the host's own switch and
	 * fires only when that changes, so a collaboration `viewer` role lock is never written back.
	 */
	@Output() readOnlyChange = new EventEmitter<boolean>();
	/** Cancelable: call `preventDefault()` on the emitted event to take a File command over. */
	@Output() fileCommand = new EventEmitter<CustomEvent<FileCommandDetail>>();
	@Output() workbookWarning = new EventEmitter<string[]>();
	@Output() ribbonAction = new EventEmitter<{ id: string }>();
	@Output() ribbonCustomize = new EventEmitter<string[]>();
	@Output() sheetChange = new EventEmitter<{ index: number; name: string }>();
	@Output() ribbonAddIn = new EventEmitter<RibbonAddInCommandDetail>();
	@Output() collaborationChange = new EventEmitter<XlsxCollaborationState>();
	@Output() ready = new EventEmitter<XlsxEditorElement>();
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
			...angularOutputs<EditorEventOptions>(EDITOR_CONTRACT, this, EDITOR_CALLBACKS),
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
