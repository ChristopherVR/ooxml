import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { fillDocxFormFields, listDocxFormFields } from './index';
import { formTemplateFixture } from '../docx/test-support/form-template-fixture';

const documentXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string')) ?? '';

describe('Word form filling from bytes', () => {
	it('lists the fields of a form and its forms protection', async () => {
		const { fields, protection } = await listDocxFormFields(
			await formTemplateFixture({ protect: true }),
		);
		expect(fields).toHaveLength(26);
		expect(protection).toEqual({ edit: 'forms', enforced: true });
		expect(fields[0]).toMatchObject({ id: 'field#1', label: 'Project', type: 'text' });
	});

	it('fills fields and returns the saved document, which lists the new values', async () => {
		const input = await formTemplateFixture();
		const filled = await fillDocxFormFields(input, {
			'field:Status': 'Withdrawn',
			'control:302': '2026-12-24',
		});
		expect(filled.results).toEqual([
			{ id: 'field:Status', ok: true, value: 'Withdrawn' },
			{ id: 'control:302', ok: true, value: '24 December 2026' },
		]);
		const { fields } = await listDocxFormFields(filled.bytes);
		expect(fields.find((field) => field.id === 'field:Status')?.value).toBe('Withdrawn');
		expect(fields.find((field) => field.id === 'control:302')?.value).toBe('24 December 2026');
		expect(await documentXml(filled.bytes)).toContain('w:fullDate="2026-12-24T00:00:00Z"');
	});

	it('writes nothing when a value is refused, unless asked to write the others', async () => {
		const input = await formTemplateFixture();
		const values = { 'field:Status': 'Closed', 'field#1': 'East' };
		await expect(fillDocxFormFields(input, values)).rejects.toThrow(
			'No field was filled: field:Status: not one of the entries: New, Revised, Withdrawn',
		);
		const partial = await fillDocxFormFields(input, values, { partial: true });
		expect(partial.results.map((result) => result.ok)).toEqual([false, true]);
		expect(await documentXml(partial.bytes)).toContain('>East<');
	});
});
