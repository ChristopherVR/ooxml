import { plainText, type Block, type DocumentModel } from 'ooxml-core/docx';
import {
	finishMailMergeIndividually,
	matchMailMergeField,
	type MailMergeData,
	type MailMergeOptions,
} from 'ooxml-core/docx/mail-merge';
import { dialogButton, labelled, selectOf, textInput } from '../dialog-fields';
import type { EditorCore } from '../editor-core';
import { saveEditorModel } from '../editor-save';
import { dialogMessage, type MailingsDialog } from './dialog';
import type { MailMergeEmailMessage, MailMergeEmailProvider } from './email-provider';
import type { MailingsLabel } from './strings';

const MAX_DRAFTS = 5;
const MAX_DRAFT_URL = 1800;
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

type EmailFormat = 'text' | 'html' | 'docx';

interface EmailSelection {
	column: string;
	subject: string;
	format: EmailFormat;
}

function address(value: unknown): string | undefined {
	if (typeof value !== 'string' || value.length > 254) return undefined;
	const trimmed = value.trim();
	return /^[A-Z0-9.!$'*+/=^_`{|}~-]+@[A-Z0-9](?:[A-Z0-9-]*[A-Z0-9])?(?:\.[A-Z0-9](?:[A-Z0-9-]*[A-Z0-9])?)+$/i.test(
		trimmed,
	)
		? trimmed
		: undefined;
}

function emailColumns(data: MailMergeData): string[] {
	return (data.fields ?? []).filter((field) =>
		data.records.some((record) => Object.hasOwn(record, field) && address(record[field])),
	);
}

function draftTemplate(model: DocumentModel): boolean {
	return (
		model.blocks.every(
			(block) =>
				block.type === 'paragraph' &&
				block.runs.every((run) => !run.image && !run.noteReference && !run.equation && !run.symbol),
		) &&
		!(
			model.footnotes?.length ||
			model.endnotes?.length ||
			model.sections?.some((section) => section.headers || section.footers)
		)
	);
}

export function emailUnavailable(core: EditorCore, data: MailMergeData): MailingsLabel | undefined {
	if (core.emailProvider === false) return 'Email merge is disabled by the host.';
	if (!emailColumns(data).length) return 'Select recipients with an email address column.';
	if (!core.emailProvider && (data.records.length > MAX_DRAFTS || !draftTemplate(core.model)))
		return 'A host email provider is required for more than five recipients or non-text documents.';
	return undefined;
}

export function emailDialog(
	dialog: MailingsDialog,
	core: EditorCore,
	data: MailMergeData,
	apply: (selection: EmailSelection) => void,
): void {
	const { t } = dialog;
	const fields = emailColumns(data);
	const column = selectOf(fields.map((name) => [name, name]));
	const matched = matchMailMergeField('EmailAddress', fields, core.model.mailMerge?.fieldMappings);
	if (matched) column.value = matched;
	const subject = textInput();
	subject.maxLength = 256;
	const provider = core.emailProvider;
	const format = selectOf([
		['text', t('Plain text')],
		...(provider
			? ([
					['html', 'HTML'],
					['docx', t('DOCX attachment')],
				] as const)
			: []),
	]);
	dialog.content.append(
		labelled(t('Email address column'), column),
		labelled(t('Subject'), subject),
		labelled(t('Email format'), format),
	);
	if (!provider) dialogMessage(dialog, 'Prepare one draft per recipient in your email app.');
	const proceed = dialogButton(t(provider ? 'Review email merge' : 'Prepare email drafts'), true);
	proceed.addEventListener('click', () => {
		const choice = format.value;
		if (choice !== 'text' && (!provider || (choice !== 'html' && choice !== 'docx'))) return;
		if (!fields.includes(column.value)) return;
		apply({ column: column.value, subject: subject.value.replace(/[\r\n]/g, ' '), format: choice });
		dialog.close();
	});
	dialog.actions(proceed);
}

const htmlEscape = (text: string) =>
	text.replace(/[&<>"']/g, (value) => {
		switch (value) {
			case '&':
				return '&amp;';
			case '<':
				return '&lt;';
			case '>':
				return '&gt;';
			case '"':
				return '&quot;';
			default:
				return '&#39;';
		}
	});

function emailHtml(blocks: readonly Block[]): string {
	return blocks
		.map((block) => {
			if (block.type === 'table')
				return `<table>${block.rows
					.map(
						(row) =>
							`<tr>${row.map((cell) => `<td>${emailHtml(cell.paragraphs)}</td>`).join('')}</tr>`,
					)
					.join('')}</table>`;
			return `<p>${block.runs
				.map((run) => {
					if (run.fieldCode || run.fieldChar) return '';
					let text = htmlEscape(run.text).replace(/\r\n?|\n/g, '<br>');
					if (run.bold) text = `<strong>${text}</strong>`;
					if (run.italic) text = `<em>${text}</em>`;
					return text;
				})
				.join('')}</p>`;
		})
		.join('');
}

export function prepareEmailMerge(
	model: DocumentModel,
	data: MailMergeData,
	selection: EmailSelection,
	options: MailMergeOptions,
) {
	const result = finishMailMergeIndividually(model, data, options);
	const jobs = result.documents.map(({ recordNumber, document }) => {
		const record = data.records[recordNumber - 1];
		const to =
			record && Object.hasOwn(record, selection.column)
				? address(record[selection.column])
				: undefined;
		return { to, document, recordNumber };
	});
	return { ...result, jobs };
}

type EmailJob = ReturnType<typeof prepareEmailMerge>['jobs'][number];

function draftUrl(job: EmailJob, subject: string): string | undefined {
	if (!job.to) return undefined;
	const encode = (value: string) =>
		encodeURIComponent(value.replace(/[\uD800-\uDFFF]/gu, '\uFFFD'));
	const [local, domain] = job.to.split('@');
	const text = plainText(job.document).replace(/\r\n?|\n/g, '\r\n');
	if (text.length > MAX_DRAFT_URL) return undefined;
	const url = `mailto:${encode(local ?? '')}@${domain}?subject=${encode(subject)}&body=${encode(text)}`;
	return url.length <= MAX_DRAFT_URL ? url : undefined;
}

export function reviewMailDrafts(
	dialog: MailingsDialog,
	jobs: readonly EmailJob[],
	subject: string,
): void {
	const drafts = jobs.map((job) => ({ to: job.to, href: draftUrl(job, subject) }));
	if (drafts.length > MAX_DRAFTS || drafts.some(({ href }) => !href)) {
		dialogMessage(dialog, 'A message is too large for an email draft. Use a host email provider.');
	} else {
		dialogMessage(dialog, 'Open each draft and send it from your email app.');
		const list = document.createElement('ul');
		for (const { to, href } of drafts) {
			if (!href || !to) continue;
			const item = document.createElement('li');
			const link = document.createElement('a');
			link.href = href;
			link.textContent = to;
			item.append(link);
			list.append(item);
		}
		dialog.content.append(list);
	}
	dialog.actions();
}

export function reviewProviderEmail(
	dialog: MailingsDialog,
	jobs: readonly EmailJob[],
	selection: EmailSelection,
	confirm: () => void,
): void {
	if (selection.format === 'docx' && jobs.length > 10) {
		dialogMessage(dialog, 'Select at most ten recipients for DOCX attachments.');
		dialog.actions();
		return;
	}
	dialogMessage(dialog, 'Send the merged messages through the host email provider?');
	const details = document.createElement('p');
	details.textContent = `${dialog.t('Subject')}: ${selection.subject}`;
	const list = document.createElement('ul');
	list.className = 'dve-mailings-table';
	for (const job of jobs) {
		const item = document.createElement('li');
		item.textContent = job.to ?? '';
		list.append(item);
	}
	dialog.content.append(details, list);
	const send = dialogButton(dialog.t('Send messages'), true);
	send.addEventListener('click', () => {
		confirm();
		dialog.close();
	});
	dialog.actions(send);
}

export async function deliverEmail(
	core: EditorCore,
	provider: MailMergeEmailProvider,
	jobs: readonly EmailJob[],
	selection: EmailSelection,
	current: () => boolean,
): Promise<{ sent: number; error?: unknown }> {
	const messages: MailMergeEmailMessage[] = [];
	let size = 0;
	let sent = 0;
	try {
		for (const job of jobs) {
			if (!current() || !job.to) return { sent };
			const base = { recipients: [job.to], subject: selection.subject };
			if (selection.format === 'docx') {
				const bytes = await saveEditorModel(core, job.document);
				size += bytes.length;
				messages.push({
					...base,
					text: '',
					attachments: [
						{ fileName: `merged-${job.recordNumber}.docx`, contentType: DOCX_TYPE, bytes },
					],
				});
			} else if (selection.format === 'html') {
				const html = emailHtml(job.document.blocks);
				size += html.length * 2;
				messages.push({ ...base, html, attachments: [] });
			} else {
				const text = plainText(job.document);
				size += text.length * 2;
				messages.push({ ...base, text, attachments: [] });
			}
			if (size > 16 * 1024 * 1024) throw new RangeError('The email batch exceeds 16 MiB.');
		}
		for (const message of messages) {
			if (!current() || core.emailProvider !== provider) break;
			await provider.send(message);
			sent++;
		}
		return { sent };
	} catch (cause) {
		return {
			sent,
			error: cause instanceof Error ? cause : new Error('Email merge failed.', { cause }),
		};
	}
}
