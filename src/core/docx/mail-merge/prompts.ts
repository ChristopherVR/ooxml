import type { DocumentModel } from '../model';
import { selectMailMergeRecipients, type MailMergeData } from './data-source';
import type { MailMergeOptions, MailMergePrompt } from './field-values';
import { projectMerge } from './project';

class PendingPrompt {
	constructor(readonly request: MailMergePrompt) {}
}

/**
 * Resolves only prompts reached by the selected records and their rules. Projection resumes by
 * replaying the current output copy with its known answers. Returning undefined cancels the run.
 */
export async function resolveMailMergePrompts(
	model: DocumentModel,
	data: MailMergeData,
	ask: (request: MailMergePrompt) => Promise<string | undefined>,
	options: Pick<MailMergeOptions, 'locale' | 'fieldMappings'> = {},
	current: () => boolean = () => true,
): Promise<MailMergeOptions | undefined> {
	const prepared = selectMailMergeRecipients(data);
	const fieldMappings = options.fieldMappings ?? model.mailMerge?.fieldMappings;
	const answers = new Map<string, string>();
	const resolved = { ...options, ...(fieldMappings && { fieldMappings }) };
	const projectOptions: MailMergeOptions = {
		...resolved,
		answerPrompt(request) {
			const answer = answers.get(request.id);
			if (answer !== undefined) return answer;
			throw new PendingPrompt(request);
		},
	};
	let sequence = 1;
	let copyPrompts = 0;
	for (let index = 0; index < prepared.records.length && current();) {
		try {
			const copy = projectMerge(model, prepared, index, sequence, projectOptions);
			index = copy.nextIndex;
			if (!copy.skipped) sequence++;
			copyPrompts = 0;
		} catch (cause) {
			if (!(cause instanceof PendingPrompt)) throw cause;
			if (++copyPrompts > 64 || answers.size >= 1000)
				throw new RangeError('Too many mail merge prompts.');
			const answer = await ask(cause.request);
			if (answer === undefined) return undefined;
			answers.set(cause.request.id, answer.slice(0, 4096));
		}
	}
	return current()
		? { ...resolved, answerPrompt: (request) => answers.get(request.id) }
		: undefined;
}
