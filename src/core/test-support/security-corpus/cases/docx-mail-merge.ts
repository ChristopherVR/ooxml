import {
	finishMailMergeIndividually,
	mailMergeField,
	mailMergeRule,
} from '../../../docx/mail-merge';
import { createDocument } from '../../../docx/model';
import { plainText } from '../../../docx/document-stats';
import type { SecurityCase } from '../case';

export const cases: readonly SecurityCase[] = [
	{
		name: 'recipient-specific output cannot contain data from the next record',
		limits: { ms: 1000, heapMb: 64 },
		run() {
			const model = createDocument();
			model.blocks = [
				{
					type: 'paragraph',
					id: 'p1',
					runs: [...mailMergeRule({ kind: 'NEXT' }), mailMergeField('Name')],
				},
			];
			const { documents, issues } = finishMailMergeIndividually(
				model,
				{ records: [{ Name: 'Ada' }, { Name: 'Bea' }] },
				{ allowRecordAdvance: false },
			);
			return {
				names: documents.map(({ document }) => plainText(document)),
				issues: issues.map(({ code }) => code),
			};
		},
		expected: {
			names: ['Ada', 'Bea'],
			issues: ['record-advance-disabled', 'record-advance-disabled'],
		},
	},
];
