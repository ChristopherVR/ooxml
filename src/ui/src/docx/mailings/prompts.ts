import type { MailMergePrompt } from 'ooxml-core/docx/mail-merge';
import { labelled, textInput } from '../dialog-fields';
import type { EditorLocale } from '../localization';
import { dialogSubmit, type MailingsDialog } from './dialog';

export type ShowMergeDialog = (build: (dialog: MailingsDialog) => void) => void;

export function askMergePrompt(
	show: ShowMergeDialog,
	locale: EditorLocale,
	request: MailMergePrompt,
): Promise<string | undefined> {
	return new Promise((resolve) => {
		show((dialog) => {
			const position = document.createElement('p');
			position.textContent = `${dialog.t('Record')} ${request.recordNumber} · ${request.kind}`;
			const input = textInput();
			input.maxLength = 1024;
			input.value = request.defaultAnswer.slice(0, input.maxLength);
			dialog.content.append(
				position,
				labelled(request.text.slice(0, 1024) || dialog.t('Prompt'), input),
			);
			const accept = () => {
				resolve(input.value);
				dialog.close();
			};
			dialog.onClose(() => resolve(undefined));
			dialogSubmit(dialog, locale, accept);
			input.addEventListener('keydown', (event) => {
				if (event.key !== 'Enter' || event.isComposing) return;
				event.preventDefault();
				accept();
			});
		});
	});
}
