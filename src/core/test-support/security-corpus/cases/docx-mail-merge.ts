import {
	finishMailMergeIndividually,
	mailMergeField,
	mailMergeRule,
	mergeFields,
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
	{
		name: 'HTML merge values keep only allow-listed content and never reference outside resources',
		limits: { ms: 2000, heapMb: 96 },
		async run() {
			const model = createDocument();
			model.blocks = [{ type: 'paragraph', id: 'p1', runs: [mailMergeField('Body')] }];
			const html = [
				'<script>x</script><style>p{background:url(http://a.test/s)}</style>',
				'<iframe src="http://a.test/f"></iframe><img src="http://a.test/i.png" onerror="x">',
				'<a href="javascript:">j</a><a href="data:text/html,<p>">d</a><a href="http://a.test/">ok</a>',
				'<p style="background:url(http://a.test/b)" onclick="x">text</p>',
				'<b>'.repeat(50_000),
			].join('');
			const { model: merged, htmlIssues } = await mergeFields(
				model,
				{ values: { Body: html }, html: ['Body'] },
				{ unlink: true },
			);
			const saved = JSON.stringify(merged);
			return {
				text: plainText(merged),
				links: [...saved.matchAll(/"href":"([^"]*)"/g)].map((match) => match[1]),
				markup: /<script|<iframe|onerror|onclick|javascript:|data:text|url\(/i.test(saved),
				issues: [...htmlIssues].sort(),
			};
		},
		expected: {
			text: 'jdok\ntext',
			links: ['http://a.test/'],
			markup: false,
			issues: ['image-refused', 'link-refused'],
		},
	},
	{
		name: 'HTML merge values trim the trailing spaces of preformatted text in linear time',
		limits: { ms: 2_000, heapMb: 128 },
		async run() {
			const model = createDocument();
			model.blocks = [{ type: 'paragraph', id: 'p1', runs: [mailMergeField('Body')] }];
			const { model: merged } = await mergeFields(
				model,
				{ values: { Body: `<pre>c${' '.repeat(150_000)}d </pre>` }, html: ['Body'] },
				{ unlink: true },
			);
			return { length: plainText(merged).length };
		},
		expected: { length: 150_002 },
	},
];
