/**
 * CommentsPanelComponent: the comments pane over the viewer's shared comments panel state.
 *
 * No Angular TestBed (see `vitest.config.ts`): the component is instantiated directly with a
 * stub TranslateService and a `CommentsPanelStateService` built on the real shared controller.
 * Pins that drafts are kept per slide, that a submit trims the draft and moves its `@`-mention
 * spans with the trim, that edits save in place, and that a read-only deck offers no edit.
 */
import { ElementRef, Injector, runInInjectionContext, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { createCommentsPanelController } from 'ooxml-ui/pptx';
import type { PptxComment, PptxCommentMention, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { CommentsPanelStateService } from './comments-panel-state.service';
import { CommentsPanelComponent } from './comments-panel.component';

function createPanel(canEdit = true) {
	let slides: PptxSlide[] = [
		{
			id: 's1',
			rId: 's1',
			slideNumber: 1,
			elements: [],
			comments: [{ id: 'c1', text: 'First', author: 'Ann' } as PptxComment],
		},
		{ id: 's2', rId: 's2', slideNumber: 2, elements: [] },
	];
	const active = signal(0);
	const controller = createCommentsPanelController({
		slides: () => slides,
		canEdit: () => canEdit,
		authorName: () => 'Me',
		commit: (next) => {
			slides = next;
		},
	});
	const state = signal(controller.getSnapshot());
	controller.subscribe(() => state.set(controller.getSnapshot()));
	const service = {
		controller,
		state: state.asReadonly(),
		slideIndex: () => active(),
		slideId: () => slides[active()]?.id,
		canEdit: () => canEdit,
	} as unknown as CommentsPanelStateService;
	const translate = { instant: (key: string) => key } as unknown as TranslateService;
	const panel = runInInjectionContext(
		Injector.create({
			providers: [
				{ provide: ElementRef, useValue: new ElementRef(document.createElement('div')) },
				{ provide: TranslateService, useValue: translate },
				{ provide: CommentsPanelStateService, useValue: service },
			],
		}),
		() => new CommentsPanelComponent(),
	);
	return { panel, active, slides: () => slides };
}

/** The component's protected template handlers, for a test without a template. */
type Handlers = {
	setDraft(text: string): void;
	setDraftMentions(mentions: PptxCommentMention[]): void;
	startEdit(id: string): void;
	setEditDraft(id: string, event: Event): void;
	saveEdit(id: string): void;
	setReplyDraft(id: string, text: string): void;
	setReplyMentions(id: string, mentions: PptxCommentMention[]): void;
	canEdit(): boolean;
};
const handlers = (panel: CommentsPanelComponent) => panel as unknown as Handlers;

const alice = (startIndex: number): PptxCommentMention => ({
	id: '{M}',
	personId: '{A}',
	authorName: 'Alice',
	startIndex,
	length: 6,
});
const covered = ({ text, mentions }: PptxComment) =>
	(mentions ?? []).map((m) => text!.slice(m.startIndex, m.startIndex + m.length));

describe('comments panel', () => {
	it('keeps the draft per slide and moves the mentions with the trimmed text', () => {
		const { panel, active, slides } = createPanel();
		handlers(panel).setDraft('  hey @Alice ');
		handlers(panel).setDraftMentions([alice(6)]);
		active.set(1);
		expect(panel.draft()).toBe('');
		active.set(0);
		expect(panel.draft()).toBe('  hey @Alice ');
		panel.submit(new Event('submit'));
		const added = slides()[0]!.comments!.at(-1)!;
		expect(added.text).toBe('hey @Alice');
		expect(covered(added)).toStrictEqual(['@Alice']);
		expect(panel.draft()).toBe('');
	});

	it('moves the mention spans with the trimmed text of a reply', () => {
		const { panel, slides } = createPanel();
		panel.startReply('c1');
		handlers(panel).setReplyDraft('c1', ' ok @Alice');
		handlers(panel).setReplyMentions('c1', [alice(4)]);
		panel.submitReply('c1');
		const reply = slides()[0]!.comments![0]!.replies![0]!;
		expect(covered(reply)).toStrictEqual(['@Alice']);
		expect(panel.replyingTo()).toBeNull();
	});

	it('edits a comment in place', () => {
		const { panel, slides } = createPanel();
		handlers(panel).startEdit('c1');
		expect(panel.editingId()).toBe('c1');
		const textarea = document.createElement('textarea');
		textarea.value = 'Rewritten';
		const event = new Event('input');
		Object.defineProperty(event, 'target', { value: textarea });
		handlers(panel).setEditDraft('c1', event);
		handlers(panel).saveEdit('c1');
		expect(slides()[0]!.comments![0]!.text).toBe('Rewritten');
		expect(panel.editingId()).toBeNull();
	});

	it('refuses every command on a read-only deck', () => {
		const { panel, slides } = createPanel(false);
		expect(handlers(panel).canEdit()).toBeFalsy();
		handlers(panel).setDraft('Nope');
		expect(panel.canAdd()).toBeFalsy();
		panel.submit(new Event('submit'));
		expect(slides()[0]!.comments).toHaveLength(1);
	});
});
