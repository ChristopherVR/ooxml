import {
	DestroyRef,
	Injector,
	runInInjectionContext,
	signal,
	ɵChangeDetectionScheduler as ChangeDetectionScheduler,
	ɵEffectScheduler as EffectScheduler,
} from '@angular/core';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { CollaborationService } from './collaboration.service';
import { EditorStateService } from './editor-state.service';
import { reportOnChange } from './report-on-change';
import { ViewerCanvasEditingService } from './viewer-canvas-editing.service';
import { ViewerDialogsService } from './viewer-dialogs.service';
import { ViewerFormatPainterService } from './viewer-format-painter.service';

/**
 * The host outputs every binding reports the same way (issue 151, item 5): the Angular wiring
 * behind `documentChange`, `notesUpdate` and the transition outputs.
 */

function slide(id: string, elements: PptxElement[] = []): PptxSlide {
	return { id, rId: id, slideNumber: 1, elements } as PptxSlide;
}

function shape(id: string): PptxElement {
	return { type: 'shape', id, x: 0, y: 0, width: 10, height: 10 } as PptxElement;
}

describe('editor edit revision (documentChange)', () => {
	it('stays put for a load and moves for each committed edit, undo and redo', () => {
		const svc = new EditorStateService();
		svc.setSlides([slide('s1', [shape('a')])]);
		expect(svc.editRevision()).toBe(0);

		svc.updateElement(0, 'a', { x: 5 });
		expect(svc.editRevision()).toBe(1);

		svc.undo();
		expect(svc.editRevision()).toBe(2);
		svc.redo();
		expect(svc.editRevision()).toBe(3);
	});

	it('does not move when undo or redo have nothing to do', () => {
		const svc = new EditorStateService();
		svc.setSlides([slide('s1', [shape('a')])]);

		svc.undo();
		svc.redo();

		expect(svc.editRevision()).toBe(0);
	});
});

describe('notesUpdate', () => {
	it('reports the new speaker-notes text after the edit lands', () => {
		const svc = new EditorStateService();
		svc.setSlides([slide('s1')]);
		const injector = Injector.create({
			providers: [
				{ provide: EditorStateService, useValue: svc },
				{ provide: ViewerDialogsService, useValue: {} },
				{ provide: ViewerFormatPainterService, useValue: {} },
				{
					provide: CollaborationService,
					useValue: { livePatcher: { flush: vi.fn(), isActive: () => false } },
				},
			],
		});
		const controller = runInInjectionContext(injector, () => new ViewerCanvasEditingService());
		const notesUpdated = vi.fn();
		controller.bind({
			canEdit: () => true,
			activeSlide: () => svc.slides()[0],
			activeSlideIndex: () => 0,
			activeTemplateElements: () => [],
			notesUpdated,
		});

		controller.onNotesUpdate('Say hello');

		expect(svc.slides()[0]?.notes).toBe('Say hello');
		expect(notesUpdated).toHaveBeenCalledExactlyOnceWith('Say hello');
	});
});

describe('reportOnChange', () => {
	it('reports transitions, never the value it was created with', () => {
		const queued = new Set<{ run(): void }>();
		const effects = {
			add: (effect: { run(): void }) => queued.add(effect),
			schedule: (effect: { run(): void }) => queued.add(effect),
			remove: (effect: { run(): void }) => queued.delete(effect),
		};
		const flush = (): void => {
			for (const effect of [...queued]) {
				queued.delete(effect);
				effect.run();
			}
		};
		const injector = Injector.create({
			providers: [
				{ provide: DestroyRef, useValue: { onDestroy: () => () => {} } },
				{ provide: ChangeDetectionScheduler, useValue: { notify: () => {} } },
				{ provide: EffectScheduler, useValue: effects },
			],
		});
		const status = signal('idle');
		const report = vi.fn();
		runInInjectionContext(injector, () => reportOnChange(status, report));

		flush();
		expect(report).not.toHaveBeenCalled();

		status.set('saving');
		flush();
		status.set('saved');
		flush();
		flush();

		expect(report.mock.calls).toStrictEqual([['saving'], ['saved']]);
	});
});
