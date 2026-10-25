/**
 * CommentsPanelComponent: the comments pane's add and reply composers, Angular binding.
 *
 * No Angular TestBed (see `vitest.config.ts`): the component is instantiated directly with a
 * stub TranslateService. Pins that a submit trims the draft and moves its `@`-mention spans
 * with the trim, so the emitted spans cover the emitted text.
 */
import { Injector, runInInjectionContext } from '@angular/core';
import type { OutputEmitterRef } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import type { PptxCommentMention } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { CommentsPanelComponent } from './comments-panel.component';
import type { CommentSubmission } from './comments-panel.component';

function createPanel(): CommentsPanelComponent {
	const translate = { instant: (key: string) => key } as unknown as TranslateService;
	return runInInjectionContext(
		Injector.create({ providers: [{ provide: TranslateService, useValue: translate }] }),
		() => new CommentsPanelComponent(),
	);
}

function captured<T>(emitter: unknown): T[] {
	const out: T[] = [];
	vi.spyOn(emitter as OutputEmitterRef<T>, 'emit').mockImplementation((value) => {
		out.push(value);
	});
	return out;
}

const alice = (startIndex: number): PptxCommentMention => ({
	id: '{M}',
	personId: '{A}',
	authorName: 'Alice',
	startIndex,
	length: 6,
});
const covered = ({ text, mentions }: CommentSubmission) =>
	mentions.map((m) => text.slice(m.startIndex, m.startIndex + m.length));

describe('comments panel submissions', () => {
	it('moves the mention spans with the trimmed text of a new comment', () => {
		const panel = createPanel();
		const added = captured<CommentSubmission>(panel.add);
		panel.draft.set('  hey @Alice ');
		panel.draftMentions.set([alice(6)]);
		panel.submit(new Event('submit'));
		expect(added).toHaveLength(1);
		expect(added[0]!.text).toBe('hey @Alice');
		expect(covered(added[0]!)).toStrictEqual(['@Alice']);
	});

	it('moves the mention spans with the trimmed text of a reply', () => {
		const panel = createPanel();
		const replied = captured<{ parentId: string } & CommentSubmission>(panel.reply);
		panel.startReply('c1');
		panel.replyDraft.set(' ok @Alice');
		panel.replyMentions.set([alice(4)]);
		panel.submitReply('c1');
		expect(replied[0]!.parentId).toBe('c1');
		expect(covered(replied[0]!)).toStrictEqual(['@Alice']);
	});
});
