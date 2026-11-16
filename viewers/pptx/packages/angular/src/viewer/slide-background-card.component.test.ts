/**
 * slide-background-card.component.test.ts: "Hide Background Graphics"
 * checkbox (`p:sld/@showMasterSp`) added to the BACKGROUND card, parity with
 * react/vue/svelte/vanilla.
 *
 * No TestBed (matching the rest of this package): the component is
 * constructed inside a plain `Injector` context, mirroring
 * `slide-transition-card.component.test.ts`.
 */
import { Injector, runInInjectionContext } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { inspectorFieldPatch, SLIDE_BACKGROUND_SECTIONS } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { SlideBackgroundCardComponent } from './slide-background-card.component';

/** The protected surface the template binds to. */
interface BackgroundCardControls {
	backgroundSections: unknown;
	patchSlide: (changes: Partial<PptxSlide>) => void;
}

function slide(overrides: Partial<PptxSlide> = {}): PptxSlide {
	return { id: 's1', rId: 's1', slideNumber: 1, elements: [], ...overrides } as PptxSlide;
}

function harness(overrides: Partial<PptxSlide> = {}): {
	editor: EditorStateService;
	controls: BackgroundCardControls;
} {
	const editor = new EditorStateService();
	editor.setSlides([slide(overrides)]);
	const injector = Injector.create({
		providers: [
			{ provide: EditorStateService, useValue: editor },
			{ provide: LoadContentService, useValue: { slideMasters: () => [], getHandler: () => null } },
			{ provide: TranslateService, useValue: { instant: (key: string) => key } },
		],
	});
	const card = runInInjectionContext(injector, () => new SlideBackgroundCardComponent());
	// See the module doc: `input()` fields are overwritten with plain functions
	// since nothing here binds them through Angular's real input machinery.
	(card as unknown as { slideIndex: () => number }).slideIndex = () => 0;
	(card as unknown as { canEdit: () => boolean }).canEdit = () => true;
	return { editor, controls: card as unknown as BackgroundCardControls };
}

describe('slide background card: shared Background fields', () => {
	const hide = (target: PptxSlide | undefined, on: boolean) =>
		inspectorFieldPatch(
			SLIDE_BACKGROUND_SECTIONS,
			'slideBackground',
			'hideBackgroundGraphics',
			target,
			on,
		)!;

	it('mounts the shared descriptors', () => {
		const { controls } = harness();
		expect(controls.backgroundSections).toBe(SLIDE_BACKGROUND_SECTIONS);
	});

	it('sets showMasterShapes: false on the slide when Hide background graphics is ticked', () => {
		const { editor, controls } = harness();
		controls.patchSlide(hide(editor.slides()[0], true));
		expect(editor.slides()[0]?.showMasterShapes).toBeFalsy();
	});

	it('sets showMasterShapes: true on the slide when it is cleared', () => {
		const { editor, controls } = harness({ showMasterShapes: false });
		controls.patchSlide(hide(editor.slides()[0], false));
		expect(editor.slides()[0]?.showMasterShapes).toBeTruthy();
	});
});
