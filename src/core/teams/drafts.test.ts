import { createDraftStore } from './drafts';

const context = { channelId: 'channel' };
const blank = { text: '', files: [], missingFiles: [] };
describe('personal message drafts', () => {
	it('restores text and attachment names without pretending the bytes survived reload', () => {
		const values = new Map<string, string>();
		const storage = {
			getItem: (key: string) => values.get(key) ?? null,
			setItem: (key: string, value: string) => {
				values.set(key, value);
			},
		};
		const drafts = createDraftStore(storage, 'ada');
		const file = Object.assign(new Blob(['xlsx']), { name: 'Budget.xlsx' });
		drafts.set(context, { text: '  unfinished\nmessage  ', files: [file], missingFiles: [] });
		expect(drafts.read(context).files[0]).toBe(file);
		expect(createDraftStore(storage, 'ada').read(context)).toEqual({
			text: '  unfinished\nmessage  ',
			files: [],
			missingFiles: ['Budget.xlsx'],
		});
		expect(createDraftStore(storage, 'bob').list()).toEqual([]);
		const snapshot = drafts.read(context);
		snapshot.files.length = 0;
		expect(drafts.read(context).files).toHaveLength(1);
	});
	it('isolates compose contexts and never restores an old send over a newer draft', () => {
		const drafts = createDraftStore(undefined, 'ada');
		const thread = { ...context, threadId: 'root', replyTo: 'child' };
		const edit = { ...thread, editId: 'child' };
		drafts.set(context, { ...blank, text: 'post' });
		drafts.set(thread, { ...blank, text: 'reply' });
		drafts.set(edit, { ...blank, text: 'edit' });
		expect(drafts.list()).toHaveLength(3);
		const revision = drafts.clear(thread);
		drafts.set(thread, { ...blank, text: 'new reply' });
		drafts.restore(thread, { ...blank, text: 'old reply' }, revision);
		expect(drafts.read(thread).text).toBe('new reply');
		expect(drafts.read(context).text).toBe('post');
		expect(drafts.read(edit).text).toBe('edit');
		const untouched = drafts.clear(edit);
		drafts.restore(edit, { ...blank, text: 'failed edit' }, untouched);
		expect(drafts.read(edit).text).toBe('failed edit');
	});
	it('ignores corrupt storage and keeps drafts usable when writes are blocked', () => {
		const drafts = createDraftStore(
			{
				getItem: () => '[null,{}, {"context":{"channelId":8},"text":"bad"}]',
				setItem: () => {
					throw new Error('blocked');
				},
			},
			'ada',
		);
		expect(drafts.list()).toEqual([]);
		drafts.set(context, { ...blank, text: 'kept' });
		expect(drafts.read(context).text).toBe('kept');
		drafts.clear(context);
		expect(drafts.list()).toEqual([]);
	});
});
