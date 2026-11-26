import { expect, it } from 'vitest';
import { plainText } from '../document-stats';
import { createDocument } from '../model';
import { mailMergeRule, previewMailMerge, resolveMailMergePrompts } from './index';

function promptDocument(count: number) {
	const rule = mailMergeRule({
		kind: 'FILLIN',
		bookmark: '',
		prompt: 'Which recipient?',
		answer: '',
	});
	return {
		...createDocument(),
		blocks: [
			{
				type: 'paragraph' as const,
				id: 'p1',
				runs: Array.from({ length: count }, () => rule).flat(),
			},
		],
	};
}

it('resolves a per-record prompt for more than a thousand recipients', async () => {
	const model = promptDocument(1);
	const data = { records: Array.from({ length: 1001 }, () => ({})) };
	let asked = 0;
	const options = await resolveMailMergePrompts(model, data, async ({ recordNumber }) => {
		asked++;
		return String(recordNumber);
	});
	expect(asked).toBe(1001);
	expect(plainText(previewMailMerge(model, data, 1001, options).document)).toBe('1001');
});

it('keeps the 64-prompt limit for a single copy even with more recipients', async () => {
	let asked = 0;
	await expect(
		resolveMailMergePrompts(promptDocument(65), { records: [{}, {}] }, async () => {
			asked++;
			return 'Answer';
		}),
	).rejects.toThrow('Too many mail merge prompts.');
	expect(asked).toBe(64);
});
