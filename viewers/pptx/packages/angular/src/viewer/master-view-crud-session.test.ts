/**
 * A Slide Master view sidebar command, run on the deck as it is now and adopted without
 * re-seeding the viewer (see `master-view-crud-session.ts`): the real handler over
 * `sample-deck.pptx` (seven slides on one master), the real loader and editor state.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { DestroyRef, Injector, runInInjectionContext } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { masterViewCrudActions } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { masterViewDeck, runMasterViewCrud } from './master-view-crud-session';
import { buildSaveSlides } from './template-mode';

const DECK = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/sample-deck.pptx');
const ON_MASTER = { tab: 'slides' as const, masterIndex: 0, layoutIndex: null };

async function open() {
	const bytes = readFileSync(DECK);
	const handler = new PptxHandler();
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	const destroyRefStub: Pick<DestroyRef, 'onDestroy'> = { onDestroy: () => () => {} };
	const injector = Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: destroyRefStub },
			{ provide: TranslateService, useValue: { instant: (key: string) => key } },
			{ provide: EditorStateService, useClass: EditorStateService },
			{ provide: LoadContentService, useClass: LoadContentService },
		],
	});
	const get = <T>(token: new (...args: never[]) => T): T =>
		runInInjectionContext(injector, () => injector.get(token));
	const editor = get(EditorStateService);
	const loader = get(LoadContentService);
	// What a load leaves behind: the handler, the slides, the masters and the parsed deck.
	loader.adoptMasterViewData(handler, data);
	editor.setSlides(loader.slides());
	const saveSlides = () => buildSaveSlides(editor.slides(), editor.templateElementsBySlideId());
	return { handler, editor, loader, saveSlides };
}

describe('slide master view commands on the deck as it is now', () => {
	it('adopts a new layout without re-seeding the viewer or dropping an unsaved edit', async () => {
		const { handler, editor, loader, saveSlides } = await open();
		const moved = editor.slides()[1]!.elements[0]!;
		editor.updateElement(1, moved.id, { x: moved.x + 40 });
		expect(editor.canUndo()).toBeTruthy();
		const slides = loader.slides();
		const layouts = loader.slideMasters()[0]?.layouts?.length ?? 0;
		const dispose = vi.spyOn(handler, 'dispose');

		const result = await runMasterViewCrud(loader, editor, {
			id: 'addLayout',
			target: ON_MASTER,
			slides: saveSlides(),
		});

		expect(result?.ok).toBeTruthy();
		if (!result?.ok) {
			return;
		}
		expect(loader.getHandler()).toBe(result.handler);
		expect(loader.slideMasters()[0]?.layouts).toHaveLength(layouts + 1);
		expect(loader.slides(), 'loader.slides re-seeds the viewer').toBe(slides);
		expect(editor.canUndo(), 'the undo history stays').toBeTruthy();
		expect(editor.dirty()).toBeTruthy();
		// The slides on screen and the undo stack still show the old handler's pictures.
		expect(dispose).not.toHaveBeenCalled();
		// The next command and a save start from the rebuilt package: it holds the edit.
		const rebuilt = result.data.slides[1]?.elements.find((element) => element.id === moved.id);
		expect(rebuilt?.x).toBe(moved.x + 40);
	});

	it('lists a layout an edit started using as in use, so it cannot be deleted', async () => {
		const { editor, loader } = await open();
		const layouts = loader.slideMasters()[0]?.layouts ?? [];
		const used = new Set(editor.slides().map((slide) => slide.layoutPath));
		const spare = layouts.findIndex((layout) => !used.has(layout.path));
		expect(spare).toBeGreaterThanOrEqual(0);
		const target = { tab: 'slides' as const, masterIndex: 0, layoutIndex: spare };
		const deletable = () =>
			masterViewCrudActions(masterViewDeck(loader, editor.slides()), target).find(
				(action) => action.id === 'deleteLayout',
			)?.enabled;
		expect(deletable()).toBeTruthy();

		editor.updateSlide(0, { layoutPath: layouts[spare]!.path });

		expect(deletable()).toBeFalsy();
	});
});
