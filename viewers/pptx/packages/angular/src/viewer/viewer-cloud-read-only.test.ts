/**
 * A cloud file the host's write policy leaves read-only cannot be edited, through the public handle
 * or the ribbon, and editing returns when an editable file (or the device deck) is opened next.
 *
 * The component's edit right is `documentView().canEdit` (the shared `resolveCloudDocumentView`)
 * through `canEdit`; this package has no TestBed, so the real `ViewerFileIOService` and
 * `EditorStateService` back a stand-in for the component and the component's own prototype methods
 * run against it, like `power-point-viewer-api.test.ts` does for the mode transitions.
 */
import { computed, DestroyRef, Injector } from '@angular/core';
import type { CloudOpenedFile, DocumentStorage } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { EditorStateService } from './editor-state.service';
import { ExportService } from './export.service';
import { LoadContentService } from './load-content.service';
import { PowerPointViewerComponent } from './power-point-viewer.component';
import { RibbonHomeSectionComponent } from './ribbon-home-section.component';
import { ViewerFileIOService } from './viewer-file-io.service';

const title: PptxElement = {
	id: 'title-0',
	type: 'text',
	text: 'Quarterly report',
	x: 67,
	y: 40,
	width: 400,
	height: 60,
};

const slide = (id: string, elements: PptxElement[] = []): PptxSlide =>
	({ id, rId: id, slideNumber: 1, elements }) as PptxSlide;

function cloudFile(canEdit: boolean): CloudOpenedFile {
	return {
		entry: { id: 'item-1', name: 'Cloud.pptx', type: 'file', kind: 'pptx' },
		kind: 'pptx',
		fileName: 'Cloud.pptx',
		bytes: new Uint8Array(),
		canEdit,
		storage: { save: vi.fn() } as DocumentStorage,
	} as unknown as CloudOpenedFile;
}

function harness(hostCanEdit = true) {
	const injector = Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: { onDestroy: () => () => {} } },
			{ provide: LoadContentService, useClass: LoadContentService },
			{ provide: ExportService, useValue: { savePresentation: vi.fn() } },
			{ provide: ViewerFileIOService, useClass: ViewerFileIOService },
		],
	});
	const fileIO = injector.get(ViewerFileIOService);
	const editor = new EditorStateService();
	editor.setSlides([slide('s1', [{ ...title }]), slide('s2')]);
	const canEdit = computed(
		() =>
			fileIO.documentView({
				storage: undefined,
				canEdit: hostCanEdit,
				fileName: undefined,
				filePath: undefined,
			}).canEdit,
	);
	// The component, reduced to what the handle methods read.
	const component = {
		canEdit,
		editor,
		activeSlideIndex: () => 0,
		displaySlides: () => editor.slides(),
	} as unknown as PowerPointViewerComponent;
	const proto = PowerPointViewerComponent.prototype as unknown as Record<
		string,
		(...args: unknown[]) => unknown
	>;
	const call = (method: string, ...args: unknown[]): unknown =>
		proto[method]!.apply(component, args);
	/** Every edit method of the public handle, tried on the first slide. */
	const tryEdits = (): void => {
		call('addSlide', 0);
		call('duplicateSlides', [0]);
		call('moveSlide', 0, 1);
		call('toggleHideSlides', [0]);
		call('updateElement', 'title-0', { x: 500 });
		call('duplicateElement', 'title-0');
		call('deleteElements', ['title-0']);
		call('deleteSlides', [1]);
	};
	return { fileIO, editor, canEdit, call, tryEdits };
}

describe('a read-only cloud file in the Angular viewer', () => {
	it('refuses every handle edit and leaves the deck untouched', () => {
		const { fileIO, editor, canEdit, tryEdits } = harness();
		expect(canEdit()).toBe(true);
		fileIO.openCloudFile(cloudFile(false));
		expect(canEdit()).toBe(false);
		const before = editor.slides();
		tryEdits();
		expect(editor.slides()).toBe(before);
		expect(editor.slides().map((s) => s.id)).toStrictEqual(['s1', 's2']);
		expect(editor.slides()[0]!.elements).toStrictEqual([title]);
		expect(editor.dirty()).toBe(false);
		expect(editor.canUndo()).toBe(false);
	});

	it('disables New Slide in the Home ribbon, the command a click would run', () => {
		const { fileIO, canEdit } = harness();
		const slidesView = (
			RibbonHomeSectionComponent.prototype as unknown as { slidesView(): unknown }
		).slidesView;
		const newSlideDisabled = (): boolean => {
			const view = slidesView.call({
				canEdit,
				layoutOptions: () => [],
				currentLayoutPath: () => undefined,
				layoutPreviews: () => ({}),
				editor: { slides: () => [slide('s1')] },
				customization: () => ({ hiddenDialogs: new Set() }),
				translation: null,
				language: () => 0,
			}) as { controls: Record<string, { disabled?: boolean }> };
			return view.controls['home.slides.newSlide']!.disabled === true;
		};
		expect(newSlideDisabled()).toBe(false);
		fileIO.openCloudFile(cloudFile(false));
		expect(newSlideDisabled()).toBe(true);
	});

	it('edits again once an editable cloud file is opened, or the device deck is loaded', () => {
		const { fileIO, editor, canEdit, call, tryEdits } = harness();
		fileIO.openCloudFile(cloudFile(false));
		tryEdits();
		expect(editor.slides()).toHaveLength(2);

		fileIO.openCloudFile(cloudFile(true));
		expect(canEdit()).toBe(true);
		call('addSlide', 0);
		expect(editor.slides()).toHaveLength(3);
		call('updateElement', 'title-0', { x: 500 });
		expect(editor.slides()[0]!.elements[0]).toMatchObject({ x: 500 });
		expect(editor.dirty()).toBe(true);

		fileIO.openCloudFile(cloudFile(false));
		fileIO.clearCloudFile();
		expect(canEdit()).toBe(true);
		call('deleteSlides', [1]);
		expect(editor.slides()).toHaveLength(2);
	});

	it('stays read-only for a read-only host even when the file is editable', () => {
		const { fileIO, editor, canEdit, tryEdits } = harness(false);
		fileIO.openCloudFile(cloudFile(true));
		expect(canEdit()).toBe(false);
		const before = editor.slides();
		tryEdits();
		expect(editor.slides()).toBe(before);
	});
});
