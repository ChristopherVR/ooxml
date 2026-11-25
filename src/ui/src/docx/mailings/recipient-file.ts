import type { MailMergeData } from 'ooxml-core/docx/mail-merge';
import { readRecipientFile } from './recipient-data';

/** Parses local data off the editor thread where module workers are available. */
export function readLocalRecipients(file: File): Promise<MailMergeData> {
	if (typeof Worker === 'undefined') return readRecipientFile(file);
	let worker: Worker;
	try {
		worker = new Worker(new URL('./mailings-recipient-worker.js', import.meta.url), {
			type: 'module',
		});
	} catch {
		return readRecipientFile(file);
	}
	return new Promise((resolve, reject) => {
		const finish = () => {
			clearTimeout(timeout);
			worker.terminate();
		};
		const timeout = setTimeout(() => {
			finish();
			reject(new Error('Reading recipients timed out.'));
		}, 15_000);
		worker.addEventListener(
			'error',
			(event) => {
				event.preventDefault();
				finish();
				// A host can block workers with CSP or compile the editor without worker support.
				void readRecipientFile(file).then(resolve, reject);
			},
			{ once: true },
		);
		worker.addEventListener(
			'message',
			(event: MessageEvent<MailMergeData | null>) => {
				finish();
				if (event.data) resolve(event.data);
				else reject(new Error('Could not read recipients.'));
			},
			{ once: true },
		);
		worker.postMessage(file);
	});
}
