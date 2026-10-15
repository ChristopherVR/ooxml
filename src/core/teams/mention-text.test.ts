import { mentionDisplayText, promoteMentions } from './mention-text';
import { type MentionTarget, resolveMentions } from './mentions';

const ada: MentionTarget = { kind: 'user', id: 'ada', name: 'Ada' };
const adaL: MentionTarget = { kind: 'user', id: 'ada2', name: 'Ada Lovelace' };
const channel: MentionTarget = { kind: 'channel', id: 'channel', name: 'channel' };

describe('promoteMentions', () => {
	it('adds the @ back in front of each picked name', () => {
		expect(promoteMentions('hey Ada please look', [ada])).toBe('hey @Ada please look');
		expect(promoteMentions('Ada Lovelace and Ada', [adaL, ada])).toBe('@Ada Lovelace and @Ada');
	});

	it('lets each pick claim one occurrence, so plain words stay plain', () => {
		expect(promoteMentions('channel: the channel is quiet', [channel])).toBe(
			'@channel: the channel is quiet',
		);
		expect(promoteMentions('Ada and Ada', [ada, ada])).toBe('@Ada and @Ada');
	});

	it('respects word boundaries and existing @ and ignores unknown picks', () => {
		expect(promoteMentions('Adam and @Ada', [ada])).toBe('Adam and @Ada');
		expect(promoteMentions('hello', [ada])).toBe('hello');
		expect(promoteMentions('hi Ada', [{ kind: 'user', id: '', name: 'Ada' }])).toBe('hi Ada');
	});

	it('round-trips with the stored mention ranges', () => {
		const stored = promoteMentions('hey Ada Lovelace, ping channel', [adaL, channel]);
		const mentions = resolveMentions(stored, [adaL, channel]);
		expect(mentions.map((m) => stored.slice(m.start, m.end))).toEqual([
			'@Ada Lovelace',
			'@channel',
		]);
		expect(mentionDisplayText({ text: stored, mentions })).toBe('hey Ada Lovelace, ping channel');
		expect(mentionDisplayText({ text: 'plain' })).toBe('plain');
	});
});
