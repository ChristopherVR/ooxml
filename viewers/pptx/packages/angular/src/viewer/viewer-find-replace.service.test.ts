/**
 * The Find and Replace bar's session, Angular binding: the service keeps the shared
 * `find-replace-session` in a signal and maps its navigation onto `goTo` and the editor's
 * selection.
 *
 * Reference binding: packages/react/src/viewer/hooks/useFindReplace.ts
 */
import { Injector, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { EditorStateService } from './editor-state.service';
import { ViewerFindReplaceService } from './viewer-find-replace.service';

function text(id: string, value: string): PptxElement {
	return {
		type: 'text',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 20,
		text: value,
		textSegments: [{ text: value, style: {} }],
	} as PptxElement;
}

/** Three matches of "foo" on two slides. */
const deck = (): PptxSlide[] => [
	{ id: 'a', rId: 'rId-a', slideNumber: 1, elements: [text('a1', 'foo bar foo')] },
	{ id: 'b', rId: 'rId-b', slideNumber: 2, elements: [text('b1', 'a foo')] },
];

function harness() {
	const slides = signal<readonly PptxSlide[]>(deck());
	const selectedIds = signal<readonly string[]>([]);
	const editor = {
		slides,
		selectedIds,
		select: vi.fn((ids: readonly string[]) => selectedIds.set([...ids])),
		applyReplacement: vi.fn((next: readonly PptxSlide[], _label?: string) => slides.set(next)),
	};
	const injector = Injector.create({
		providers: [
			{ provide: EditorStateService, useValue: editor as unknown as EditorStateService },
			{
				provide: TranslateService,
				useValue: { instant: (key: string) => key } as unknown as TranslateService,
			},
			{ provide: ViewerFindReplaceService, useClass: ViewerFindReplaceService, deps: [] },
		],
	});
	const service = injector.get(ViewerFindReplaceService);
	const goTo = vi.fn<(index: number) => void>();
	service.bind(goTo);
	return { service, editor, goTo, slides, selectedIds };
}

const find = (query: string) => ({ query, matchCase: false });
const replace = (query: string, replacement: string) => ({ query, replacement, matchCase: false });
const texts = (slides: readonly PptxSlide[]) =>
	slides.map((s) => (s.elements[0] as { text?: string }).text);

describe('viewerFindReplaceService', () => {
	it('a match goes to its slide and selects its element', () => {
		const h = harness();
		h.service.openFindReplace();
		h.service.onFind(find('foo'));
		expect(h.goTo).toHaveBeenLastCalledWith(0);
		expect(h.selectedIds()).toStrictEqual(['a1']);
		h.service.onNavigate(1);
		h.service.onNavigate(1);
		expect(h.goTo).toHaveBeenLastCalledWith(1);
		expect(h.selectedIds()).toStrictEqual(['b1']);
		h.service.onNavigate(1);
		expect(h.service.activeIndex()).toBe(0);
	});

	it('closing drops the matches; reopening searches the kept query from the first', () => {
		const h = harness();
		h.service.openFindReplace();
		h.service.onFind(find('foo'));
		h.service.onNavigate(1);
		h.service.closeFindReplace();
		expect(h.service.showFindReplace()).toBeFalsy();
		expect(h.service.results()).toHaveLength(0);
		expect(h.service.activeIndex()).toBe(-1);
		h.service.openFindReplace();
		expect(h.service.session().query).toBe('foo');
		expect(h.service.results()).toHaveLength(3);
		expect(h.service.activeIndex()).toBe(0);
	});

	it('replace changes the current match only and moves on to the next one', () => {
		const h = harness();
		h.service.openFindReplace();
		h.service.onFind(find('foo'));
		h.service.onNavigate(1);
		h.service.onReplaceOne(replace('foo', 'qux'));
		expect(h.editor.applyReplacement).toHaveBeenCalledExactlyOnceWith(
			expect.any(Array),
			'pptx.undoAction.replace',
		);
		expect(texts(h.slides())).toStrictEqual(['foo bar qux', 'a foo']);
		expect(h.service.results()).toHaveLength(2);
		expect(h.service.activeIndex()).toBe(1);
		expect(h.selectedIds()).toStrictEqual(['b1']);
	});

	it('replace all applies every match as one undo step', () => {
		const h = harness();
		h.service.openFindReplace();
		h.service.onFind(find('foo'));
		h.service.onReplaceAll(replace('foo', 'qux'));
		expect(h.editor.applyReplacement).toHaveBeenCalledExactlyOnceWith(
			expect.any(Array),
			'pptx.undoAction.replaceAll',
		);
		expect(texts(h.slides())).toStrictEqual(['qux bar qux', 'a qux']);
		expect(h.service.results()).toHaveLength(0);
	});
});
