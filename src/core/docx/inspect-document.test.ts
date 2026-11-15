import { describe, expect, it } from 'vitest';
import { loadDocx } from './parse';
import { createDocument } from './model';
import { inspectDocument } from './inspect-document';
import { AUTHOR, inspectFixture, REVIEWER, TEMPLATE_PATH } from './test-support/inspect-fixture';

const byCategory = (model: Parameters<typeof inspectDocument>[0]) =>
	Object.fromEntries(inspectDocument(model).map((entry) => [entry.category, entry.findings]));

describe('inspectDocument', () => {
	it('reports nothing for a new document', () => {
		expect(inspectDocument(createDocument()).every((entry) => !entry.findings.length)).toBe(true);
	});

	it("lists each category's findings in Word's order", async () => {
		const { model } = await loadDocx(await inspectFixture());
		expect(inspectDocument(model).map((entry) => entry.category)).toEqual([
			'comments',
			'properties',
			'customXml',
			'headersFooters',
			'hiddenText',
		]);
		const found = byCategory(model);
		expect(found.comments).toEqual([
			{ kind: 'comments', count: 1 },
			{ kind: 'revisions', count: 1 },
		]);
		expect(found.properties).toEqual([
			{ kind: 'documentProperties', count: 3 },
			{ kind: 'author', count: 2, values: [AUTHOR, REVIEWER] },
			{ kind: 'reviewers', count: 2, values: [REVIEWER, AUTHOR] },
			{ kind: 'company', count: 1, values: ['Example Co'] },
			{ kind: 'manager', count: 1, values: ['Carol Manager'] },
			{ kind: 'template', count: 1, values: [TEMPLATE_PATH] },
			{ kind: 'customProperties', count: 1, values: ['Project'] },
		]);
		expect(found.customXml).toEqual([{ kind: 'customXml', count: 1 }]);
		expect(found.headersFooters).toEqual([
			{ kind: 'headers', count: 1 },
			{ kind: 'footers', count: 1 },
		]);
		expect(found.hiddenText).toEqual([{ kind: 'hiddenText', count: 1 }]);
	});

	it('inspects only the categories asked for', async () => {
		const { model } = await loadDocx(await inspectFixture());
		expect(inspectDocument(model, ['hiddenText'])).toEqual([
			{ category: 'hiddenText', findings: [{ kind: 'hiddenText', count: 1 }] },
		]);
	});

	it('counts watermarks in headers and ignores the default template', () => {
		const model = createDocument();
		model.packageInformation = { template: 'Normal.dotm' };
		model.sections = [
			{
				pageWidthTwips: 12240,
				pageHeightTwips: 15840,
				orientation: 'portrait',
				headers: {
					default: {
						blocks: [
							{
								type: 'paragraph',
								id: 'h1',
								runs: [
									{
										text: '',
										image: {
											relId: '',
											partName: '',
											contentType: 'application/octet-stream',
											widthPx: 0,
											heightPx: 0,
											watermark: {
												text: 'DRAFT',
												color: '#c0c0c0',
												semitransparent: true,
												layout: 'diagonal',
											},
										},
									},
								],
							},
						],
					},
				},
			} as never,
		];
		const found = byCategory(model);
		expect(found.headersFooters).toEqual([
			{ kind: 'headers', count: 1 },
			{ kind: 'watermarks', count: 1 },
		]);
		expect(found.properties).toEqual([]);
	});
});
