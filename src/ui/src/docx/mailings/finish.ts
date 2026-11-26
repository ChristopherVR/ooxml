import {
	finishMailMerge,
	resolveMailMergePrompts,
	type MailMergeData,
} from 'ooxml-core/docx/mail-merge';
import type { DocumentModel } from 'ooxml-core/docx';
import type { EditorCore } from '../editor-core';
import { saveEditorModel } from '../editor-save';
import { openMergedDocument } from '../file-commands';
import type { LocalizationKey } from '../localization';
import { dialogMessage, type MailingsDialog } from './dialog';
import {
	deliverEmail,
	emailDialog,
	prepareEmailMerge,
	reviewMailDrafts,
	reviewProviderEmail,
} from './email';
import { askMergePrompt } from './prompts';
import type { MailingsLabel } from './strings';
import { checkDialog, type MergeDestination } from './workflow-dialogs';

type ShowDialog = (title: LocalizationKey, build: (dialog: MailingsDialog) => void) => void;
type EmailSelection = Parameters<typeof prepareEmailMerge>[2];

export async function finishWorkflow(
	core: EditorCore,
	recipients: MailMergeData,
	destination: MergeDestination,
	show: ShowDialog,
	current: () => boolean,
	print: (document: DocumentModel) => void,
): Promise<void> {
	const title = 'Finish & Merge';
	const message = (text: MailingsLabel) =>
		show(title, (dialog) => {
			dialogMessage(dialog, text);
			dialog.actions();
		});
	const model = core.model;
	const provider = core.emailProvider;
	const locale = core.locale;
	const email =
		destination === 'email'
			? await new Promise<EmailSelection | undefined>((resolve) =>
					show(title, (dialog) => {
						dialog.onClose(() => resolve(undefined));
						emailDialog(dialog, core, recipients, resolve);
					}),
				)
			: undefined;
	if ((destination === 'email' && !email) || !current()) return;
	const options = await resolveMailMergePrompts(
		model,
		recipients,
		(request) => askMergePrompt((build) => show(title, build), locale, request),
		{ locale, ...(email && { allowRecordAdvance: false }) },
		current,
	);
	if (!options || !current()) return;
	if (email) {
		const { issues, jobs } = prepareEmailMerge(model, recipients, email, options);
		if (issues.length) show('Check for Errors', (dialog) => checkDialog(dialog, issues));
		else if (!jobs.length)
			message('No output was produced because the rules skipped every recipient.');
		else if (jobs.some(({ to }) => !to))
			message('A recipient has a missing or invalid email address.');
		else if (provider !== core.emailProvider || provider === false) return;
		else if (!provider) show(title, (dialog) => reviewMailDrafts(dialog, jobs, email.subject));
		else {
			const confirmed = await new Promise<boolean>((resolve) =>
				show(title, (dialog) => {
					dialog.onClose(() => resolve(false));
					reviewProviderEmail(dialog, jobs, email, () => resolve(true));
				}),
			);
			if (!confirmed || !current()) return;
			const { sent, error } = await deliverEmail(core, provider, jobs, email, current);
			if (!current()) return;
			show(title, (dialog) => {
				const status = dialogMessage(dialog, 'Messages accepted by the provider');
				status.append(` ${sent} / ${jobs.length}`);
				if (error)
					dialogMessage(
						dialog,
						'Email delivery stopped. Check with your provider before retrying.',
					);
				dialog.actions();
			});
			if (error) core.host.reportError(error);
		}
		return;
	}
	const result = finishMailMerge(model, recipients, options);
	if (result.issues.length)
		show('Check for Errors', (dialog) => checkDialog(dialog, result.issues));
	else if (!result.document)
		message('No output was produced because the rules skipped every recipient.');
	else if (destination === 'print') print(result.document);
	else {
		const name = (core.shell.chrome?.fileName ?? 'Document').replace(/\.[^./\\]+$/, '');
		const bytes = await saveEditorModel(core, result.document);
		if (current()) openMergedDocument(core.element, bytes, `${name}-merged.docx`);
	}
}
