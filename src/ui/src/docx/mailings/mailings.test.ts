// @vitest-environment jsdom
import { afterEach, beforeAll, beforeEach, expect, it, vi } from 'vitest';
import { createDocument, loadDocx, type DocumentModel } from 'ooxml-core/docx';
import { loadMailMergeStationery, resolveMailMergePrompts } from 'ooxml-core/docx/mail-merge';
import '../component';
import { findLocalizedControl } from '../localization';
import { must, paragraphAt, ribbonTab } from '../test-support';
import type { MailMergeEmailMessage } from './email-provider';

const execCommand = Object.getOwnPropertyDescriptor(document, 'execCommand');
beforeAll(async () => {
	await Promise.all([
		import('./controller'),
		loadMailMergeStationery(),
		resolveMailMergePrompts(createDocument(), { records: [] }, async () => undefined),
	]);
}, 120_000);
beforeEach(() => {
	Object.defineProperty(document, 'execCommand', { configurable: true, value: () => false });
});
afterEach(() => {
	document.body.replaceChildren();
	if (execCommand) Object.defineProperty(document, 'execCommand', execCommand);
	else Reflect.deleteProperty(document, 'execCommand');
	vi.restoreAllMocks();
});

async function mount(model = createDocument()) {
	const editor = document.createElement('docx-editor');
	editor.documentModel = model;
	document.body.append(editor);
	const root = must(editor.shadowRoot);
	const panel = must(root.querySelector('[data-panel="Mailings"]'));
	(await ribbonTab(root, 'mailings')).click();
	const errors: Error[] = [];
	editor.addEventListener('document-error', (event) => errors.push(event.detail));
	const ribbon = (label: string) => must(findLocalizedControl<HTMLButtonElement>(panel, label));
	const dialog = async () => {
		await vi.waitFor(() => expect(root.querySelector('.dve-mailings-dialog')).not.toBeNull(), {
			timeout: 10_000,
		});
		return must(root.querySelector('.dve-mailings-dialog'));
	};
	const button = (scope: ParentNode, label: string) =>
		must([...scope.querySelectorAll('button')].find((item) => item.textContent === label));
	const previewContains = async (text: string) =>
		vi.waitFor(() => {
			expect(errors).toEqual([]);
			expect(root.querySelector('[data-merge-projection]')?.textContent ?? '').toContain(text);
		});
	return { editor, root, ribbon, dialog, button, previewContains };
}

const content = (model: DocumentModel) =>
	model.blocks.map((_, index) =>
		paragraphAt(model.blocks, index)
			.runs.map((run) => run.text)
			.join(''),
	);

it('round-trips an edited template and emits a separate finished document after preview', async () => {
	const { editor, ribbon, dialog, button, root, previewContains } = await mount();
	const read = vi.fn(() => ({
		fields: ['Name'],
		records: [{ Name: 'Ada' }, { Name: '<b>Bea</b>' }],
	}));
	editor.mailMergeDataSource = { read };
	ribbon('Start Mail Merge').click();
	must((await dialog()).querySelector('select')).value = 'formLetters';
	button(await dialog(), 'OK').click();
	ribbon('Select Recipients').click();
	expect(read).not.toHaveBeenCalled();
	button(await dialog(), 'Use host recipients').click();
	await vi.waitFor(() => expect(ribbon('Insert Merge Field').disabled).toBe(false));
	expect(read).toHaveBeenCalledTimes(1);
	ribbon('Insert Merge Field').click();
	button(await dialog(), 'OK').click();
	expect(content(must(editor.documentModel))).toEqual(['«Name»']);
	expect(editor.getToolbarProps().onUndo()).toBe(true);
	expect(content(must(editor.documentModel))).toEqual(['']);
	expect(editor.getToolbarProps().onRedo()).toBe(true);
	expect(content(must(editor.documentModel))).toEqual(['«Name»']);
	ribbon('Highlight Merge Fields').click();
	await vi.waitFor(() => expect(editor.hasAttribute('data-highlight-merge-fields')).toBe(true));
	ribbon('Preview Results').click();
	await previewContains('Ada');
	ribbon('Last recipient').click();
	await previewContains('<b>Bea</b>');
	expect(root.querySelector('[data-merge-projection] b')).toBeNull();
	expect(ribbon('Next recipient').disabled).toBe(true);
	const template = await editor.saveBytes();
	const original = await loadDocx(template);
	expect(content(original.model)).toEqual(['«Name»']);
	expect(original.model.mailMerge?.mainDocumentType).toBe('formLetters');
	expect(paragraphAt(original.model.blocks, 0).runs[0]?.field?.instr).toBe('MERGEFIELD "Name"');
	let output: Uint8Array | undefined;
	editor.addEventListener('file-command', (event) => {
		if (event.detail.command !== 'openMergedDocument') return;
		event.preventDefault();
		output = event.detail.bytes;
	});
	ribbon('Finish & Merge').click();
	expect(button(await dialog(), 'Send Email Messages').disabled).toBe(true);
	button(await dialog(), 'Edit Individual Documents').click();
	await vi.waitFor(() => expect(output).toBeDefined());
	const finished = await loadDocx(must(output));
	expect(finished.model.mailMerge).toBeUndefined();
	expect(content(finished.model)).toEqual(['Ada', '<b>Bea</b>']);
	expect(content(must(editor.documentModel))).toEqual(['«Name»']);
	await editor.load(template);
	expect(content(must(editor.documentModel))).toEqual(['«Name»']);
	expect(ribbon('Preview Results').disabled).toBe(true);
});

it('cancels selection edits, then applies filtering, sorting and stable record exclusions', async () => {
	const { editor, dialog, button, ribbon, previewContains } = await mount();
	ribbon('Select Recipients').click();
	const unavailable = button(await dialog(), 'Use host recipients');
	expect(unavailable.disabled).toBe(true);
	expect(unavailable.title).toBe('The host has not registered a recipient provider.');
	button(await dialog(), 'Close').click();
	editor.mailMergeDataSource = {
		read: () => ({
			records: [
				{ Name: 'Zoe', Status: 'Yes' },
				{ Name: 'Ada', Status: 'Yes' },
				{ Name: 'Bea', Status: 'No' },
				{ Name: 'Cal', Status: 'Yes' },
			],
		}),
	};
	ribbon('Select Recipients').click();
	button(await dialog(), 'Use host recipients').click();
	await vi.waitFor(() => expect(ribbon('Edit Recipient List').disabled).toBe(false));
	ribbon('Edit Recipient List').click();
	must(
		(await dialog()).querySelector<HTMLInputElement>('[aria-label="Include: Record 2"]'),
	).click();
	button(await dialog(), 'Cancel').click();
	ribbon('Edit Recipient List').click();
	const scope = await dialog();
	const include = must(scope.querySelector<HTMLInputElement>('[aria-label="Include: Record 2"]'));
	expect(include.checked).toBe(true);
	include.click();
	const filter = must(scope.querySelector<HTMLInputElement>('[aria-label="Filter recipients"]'));
	filter.value = 'yes';
	const column = must(scope.querySelector<HTMLSelectElement>('[aria-label="Filter column"]'));
	column.value = 'Status';
	column.dispatchEvent(new Event('change'));
	const sort = must(scope.querySelector<HTMLSelectElement>('[aria-label="Sort by"]'));
	sort.value = 'Name';
	sort.dispatchEvent(new Event('change'));
	expect([...scope.querySelectorAll('tbody tr')].map((item) => item.textContent)).toEqual([
		'2AdaYes',
		'4CalYes',
		'1ZoeYes',
	]);
	button(scope, 'OK').click();
	ribbon('Insert Merge Field').click();
	button(await dialog(), 'OK').click();
	ribbon('Preview Results').click();
	await previewContains('Cal');
	ribbon('Next recipient').click();
	await previewContains('Zoe');
	expect(ribbon('Next recipient').disabled).toBe(true);
	editor.readOnly = true;
	expect(editor.shadowRoot?.querySelector('[data-merge-projection]')).toBeNull();
	expect(ribbon('Preview Results').disabled).toBe(true);
});

it('reports missing recipients and refuses a provider response after replacing the template', async () => {
	const { editor, dialog, button, ribbon } = await mount();
	ribbon('Check for Errors').click();
	expect((await dialog()).textContent).toContain('No recipients are included.');
	button(await dialog(), 'Close').click();
	let provide: ((data: { records: { Name: string }[] }) => void) | undefined;
	const pending = new Promise<{ records: { Name: string }[] }>((resolve) => (provide = resolve));
	editor.mailMergeDataSource = { read: () => pending };
	ribbon('Select Recipients').click();
	button(await dialog(), 'Use host recipients').click();
	editor.documentModel = createDocument();
	must(provide)({ records: [{ Name: 'old' }] });
	await pending;
	await Promise.resolve();
	expect(ribbon('Preview Results').disabled).toBe(true);
	expect(ribbon('Select Recipients').disabled).toBe(false);
});

it('inserts mapped address and greeting fields and authors IF through the Rules dialog', async () => {
	const { editor, dialog, button, ribbon, previewContains } = await mount();
	editor.mailMergeDataSource = { read: () => ({ records: [{ Given: 'Ada', City: 'London' }] }) };
	ribbon('Select Recipients').click();
	button(await dialog(), 'Use host recipients').click();
	await vi.waitFor(() => expect(ribbon('Match Fields').disabled).toBe(false));
	ribbon('Match Fields').click();
	must((await dialog()).querySelector<HTMLSelectElement>('[aria-label="First name"]')).value =
		'Given';
	button(await dialog(), 'OK').click();
	ribbon('Greeting Line').click();
	button(await dialog(), 'OK').click();
	ribbon('Address Block').click();
	button(await dialog(), 'OK').click();
	ribbon('Rules').click();
	const rule = await dialog();
	must(rule.querySelector<HTMLInputElement>('[aria-label="Compare with"]')).value = 'Ada';
	must(rule.querySelector<HTMLInputElement>('[aria-label="Then insert"]')).value = 'Matched';
	button(rule, 'OK').click();
	ribbon('Preview Results').click();
	await previewContains('Dear Ada,');
	await previewContains('LondonMatched');
	ribbon('Preview Results').click();
	await vi.waitFor(() => expect(ribbon('Start Mail Merge').disabled).toBe(false));
	ribbon('Start Mail Merge').click();
	must((await dialog()).querySelector('select')).value = '';
	button(await dialog(), 'OK').click();
	expect(must(editor.documentModel).mailMerge).toBeUndefined();
	expect(ribbon('Preview Results').disabled).toBe(true);
});

it('persists and propagates the label layout through ribbon edits', async () => {
	const { editor, ribbon, dialog, button } = await mount();
	expect(ribbon('Update Labels').disabled).toBe(true);
	ribbon('Start Mail Merge').click();
	must((await dialog()).querySelector('select')).value = 'mailingLabels';
	button(await dialog(), 'OK').click();
	button(await dialog(), 'Create merge template').click();
	expect(editor.documentModel?.mailMerge?.mainDocumentType).toBe('mailingLabels');
	editor.mailMergeDataSource = {
		read: () => ({ records: [{ FirstName: 'Ada' }, { FirstName: 'Bea' }] }),
	};
	ribbon('Select Recipients').click();
	button(await dialog(), 'Use host recipients').click();
	await vi.waitFor(() => expect(ribbon('Insert Merge Field').disabled).toBe(false));
	ribbon('Insert Merge Field').click();
	button(await dialog(), 'OK').click();
	ribbon('Update Labels').click();
	await vi.waitFor(() =>
		expect(editor.shadowRoot?.querySelectorAll('.ProseMirror td')[2]?.textContent).toContain(
			'«NEXT»',
		),
	);
	const { model } = await loadDocx(await editor.saveBytes());
	expect(model.blocks[0]).toMatchObject({
		type: 'table',
		layout: 'fixed',
		grid: [3780, 180, 3780, 180, 3780],
	});
	const table = model.blocks[0];
	if (table?.type !== 'table') throw new Error('Missing label table');
	expect(
		table.rows[0]?.[2]?.paragraphs[0]?.runs.some(({ fieldCode }) => fieldCode?.trim() === 'NEXT'),
	).toBe(true);
});

it('cancels a prompt safely and sends escaped personalized HTML only after provider confirmation', async () => {
	const model = createDocument();
	const { mailMergeRule } = await import('ooxml-core/docx/mail-merge');
	model.blocks[0] = {
		type: 'paragraph',
		id: 'p1',
		runs: mailMergeRule({
			kind: 'FILLIN',
			bookmark: '',
			prompt: 'Which line?',
			answer: 'Default',
			once: true,
		}),
	};
	const { editor, ribbon, dialog, button } = await mount(model);
	const send = vi.fn<(message: MailMergeEmailMessage) => void>();
	editor.emailProvider = { send };
	editor.mailMergeDataSource = {
		read: () => ({ records: [{ Email: 'ada@example.test' }, { Email: 'bea@example.test' }] }),
	};
	ribbon('Select Recipients').click();
	button(await dialog(), 'Use host recipients').click();
	await vi.waitFor(() => expect(ribbon('Finish & Merge').disabled).toBe(false));
	const prepare = async () => {
		ribbon('Finish & Merge').click();
		button(await dialog(), 'Send Email Messages').click();
		const email = await dialog();
		must(email.querySelector<HTMLInputElement>('[aria-label="Subject"]')).value = 'A subject';
		must(email.querySelector<HTMLSelectElement>('[aria-label="Email format"]')).value = 'html';
		button(email, 'Review email merge').click();
		return dialog();
	};
	button(await prepare(), 'Cancel').click();
	await vi.waitFor(() => expect(ribbon('Finish & Merge').disabled).toBe(false));
	expect(send).not.toHaveBeenCalled();
	const prompt = await prepare();
	const answer = must(prompt.querySelector<HTMLInputElement>('[aria-label="Which line?"]'));
	expect(answer.value).toBe('Default');
	answer.value = '<img onload="test">';
	button(prompt, 'OK').click();
	const review = await dialog();
	expect(send).not.toHaveBeenCalled();
	button(review, 'Send messages').click();
	await vi.waitFor(() => expect(send).toHaveBeenCalledTimes(2));
	expect(send.mock.calls[0]?.[0]).toEqual({
		recipients: ['ada@example.test'],
		subject: 'A subject',
		html: '<p>&lt;img onload=&quot;test&quot;&gt;</p>',
		attachments: [],
	});
	expect(send.mock.calls[1]?.[0]).toMatchObject({ recipients: ['bea@example.test'] });
});

it('blocks email output when NEXTIF would merge another recipient into the message', async () => {
	const model = createDocument();
	const { mailMergeField, mailMergeRule } = await import('ooxml-core/docx/mail-merge');
	model.blocks[0] = {
		type: 'paragraph',
		id: 'p1',
		runs: [
			...mailMergeRule({
				kind: 'NEXTIF',
				field: 'Name',
				operator: '=',
				value: 'Ada',
				yes: '',
				no: '',
			}),
			mailMergeField('Name'),
		],
	};
	const { editor, ribbon, dialog, button } = await mount(model);
	const send = vi.fn<(message: MailMergeEmailMessage) => void>();
	editor.emailProvider = { send };
	editor.mailMergeDataSource = {
		read: () => ({
			records: [
				{ Name: 'Ada', Email: 'ada@example.test' },
				{ Name: 'Bea', Email: 'bea@example.test' },
			],
		}),
	};
	ribbon('Select Recipients').click();
	button(await dialog(), 'Use host recipients').click();
	await vi.waitFor(() => expect(ribbon('Finish & Merge').disabled).toBe(false));
	ribbon('Finish & Merge').click();
	button(await dialog(), 'Send Email Messages').click();
	button(await dialog(), 'Review email merge').click();
	expect((await dialog()).textContent).toContain(
		'An email merge cannot advance to another recipient. · NEXTIF',
	);
	expect(send).not.toHaveBeenCalled();
});
