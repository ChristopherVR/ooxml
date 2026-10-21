import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { listFormFields, type FormFieldInfo } from './form-fill-fields';
import { formTemplateFixture } from './test-support/form-template-fixture';

const fields = async () => listFormFields((await loadDocx(await formTemplateFixture())).model);
const byLabel = (list: FormFieldInfo[], label: string) => {
	const found = list.find((field) => field.label === label);
	if (!found) throw new Error(`no field labelled ${label}: ${list.map((f) => f.label).join(', ')}`);
	return found;
};

describe('listing the fill-in fields of a form', () => {
	it('lists legacy form fields and content controls in reading order', async () => {
		const list = await fields();
		expect(list.map((field) => field.label)).toEqual([
			'Project',
			'Project: Date',
			'Document No.',
			'Document No.: Rev',
			'To 1',
			'To 2',
			'To 3',
			'To 4',
			'Preparer',
			'Preparer: Date',
			'Reviewer (optional)',
			'Reviewer (optional): Date',
			'Status',
			'Approved',
			'Approved: Count',
			'Archived',
			'Needs further checking: Yes',
			'Needs further checking: No',
			'Status of information',
			'Purpose of issue',
			'Kind 1',
			'Kind 2',
			'Client',
			'Due',
			'Region',
			'Owner',
		]);
	});

	it('gives each field an id from its name or control id, else from its place', async () => {
		const list = await fields();
		expect(list.map((field) => field.id)).toEqual([
			'field#1',
			'field:IssueDate',
			'field#3',
			'field#4',
			'field#5',
			'field#6',
			'field#7',
			'field#8',
			'field#9',
			'field#10',
			'field#11',
			'field#12',
			'field:Status',
			'field:Approved',
			'field:Count',
			'field:Archived',
			'control:101',
			'control:102',
			'field#17',
			'field#18',
			'control:201',
			'control:202',
			'control:301',
			'control:302',
			'control:303',
			'control:304',
		]);
		// Reading the same file again gives the same ids.
		expect((await fields()).map((field) => field.id)).toEqual(list.map((field) => field.id));
	});

	it('reads the type, constraints and current value of each kind of field', async () => {
		const list = await fields();
		expect(byLabel(list, 'Project')).toMatchObject({
			source: 'legacy',
			type: 'text',
			value: 'Project',
			defaultValue: 'Project',
			fillable: true,
			story: 'body',
		});
		expect(byLabel(list, 'Project: Date')).toMatchObject({
			type: 'date',
			value: '',
			format: 'dd-MMM-yyyy',
			name: 'IssueDate',
		});
		expect(byLabel(list, 'Document No.')).toMatchObject({ type: 'text', maxLength: 12 });
		expect(byLabel(list, 'Document No.: Rev')).toMatchObject({ format: 'UPPERCASE' });
		expect(byLabel(list, 'Status')).toMatchObject({
			type: 'dropdown',
			value: 'New',
			choices: [
				{ display: 'New', value: 'New' },
				{ display: 'Revised', value: 'Revised' },
				{ display: 'Withdrawn', value: 'Withdrawn' },
			],
		});
		expect(byLabel(list, 'Approved')).toMatchObject({ type: 'checkbox', value: false });
		expect(byLabel(list, 'Approved: Count')).toMatchObject({
			type: 'number',
			format: '#,##0.00',
			value: '',
		});
		expect(byLabel(list, 'Archived')).toMatchObject({ fillable: false, reason: 'disabled' });
		expect(byLabel(list, 'Needs further checking: Yes')).toMatchObject({
			source: 'control',
			type: 'checkbox',
			value: false,
		});
		expect(byLabel(list, 'Kind 1')).toMatchObject({
			type: 'dropdown',
			value: '',
			placeholder: 'Choose an item.',
			choices: [
				{ display: 'Reference', value: 'Reference' },
				{ display: 'Cross reference', value: 'Cross reference' },
			],
		});
		expect(byLabel(list, 'Client')).toMatchObject({
			type: 'text',
			value: '',
			alias: 'Client',
			tag: 'client',
			placeholder: 'Click or tap here to enter text.',
		});
		expect(byLabel(list, 'Due')).toMatchObject({
			type: 'date',
			format: 'dd MMMM yyyy',
			locale: 'en-US',
			value: '',
		});
		expect(byLabel(list, 'Region')).toMatchObject({
			type: 'dropdown',
			value: 'North',
			allowsOtherText: true,
		});
		expect(byLabel(list, 'Owner')).toMatchObject({ fillable: false, reason: 'locked' });
	});

	it('builds labels from the row, the column header, the text beside the field and ordinals', async () => {
		const list = await fields();
		expect(byLabel(list, 'Preparer: Date').labelParts).toEqual(['Preparer', 'Date']);
		expect(byLabel(list, 'To 3').labelParts).toEqual(['To']);
		expect(byLabel(list, 'Needs further checking: No').labelParts).toEqual([
			'Needs further checking',
			'No',
		]);
		expect(byLabel(list, 'Kind 2').labelParts).toEqual(['Kind']);
	});

	it('uses a unique tag when a control has no id of its own', async () => {
		const body =
			'<w:p><w:r><w:t xml:space="preserve">Code: </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="code"/><w:text/></w:sdtPr>' +
			'<w:sdtContent><w:r><w:t>A1</w:t></w:r></w:sdtContent></w:sdt></w:p>' +
			'<w:p><w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>B2</w:t></w:r></w:sdtContent></w:sdt></w:p>';
		const { model } = await loadDocx(await formTemplateFixture({ body }));
		const list = listFormFields(model);
		// A paragraph that holds a field is not a label: the second control is named by its type.
		expect(list.map((field) => [field.id, field.label, field.value])).toEqual([
			['tag:code', 'Code', 'A1'],
			['control#2', 'Text', 'B2'],
		]);
	});
});
