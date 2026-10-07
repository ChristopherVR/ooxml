import { createClientId } from '../collab/identity';
import { contentUrl } from './content';
import { checkFileAbort, withFileAbort, type FileOperationOptions } from './file-transfer';
import { sanitizeAttachment, sanitizeChannelName, type Attachment } from './model';
import type { UploadableFile } from './store';

export const storageFileName = (raw: string): string =>
	sanitizeChannelName(raw).replace(/[\/\\\u0000-\u001f\u007f]/gu, '');
export function validateStorageFile(file: UploadableFile): void {
	if (!storageFileName(file.name) || file.size > 33_554_432)
		throw new Error('This file cannot be saved to the channel');
}

/** A display name stays stable while every storage attempt receives a fresh name. */
export async function storeAttachment(
	file: UploadableFile & Blob,
	prefix: string,
	upload: (file: UploadableFile & Blob, signal?: AbortSignal) => Promise<Attachment>,
	options?: FileOperationOptions,
): Promise<Attachment> {
	validateStorageFile(file);
	const name = storageFileName(file.name);
	const ext = name.includes('.') ? `.${name.split('.').pop()}` : '';
	const unique = Object.assign(new Blob([file], { type: file.type ?? '' }), {
		name: `${prefix}-${createClientId()}${ext}`,
	});
	let uploaded: Attachment;
	try {
		uploaded = await withFileAbort(() => upload(unique, options?.signal), options?.signal);
	} catch (error) {
		if (error instanceof Error && error.name === 'AbortError') throw error;
		throw new Error(
			`The file could not be uploaded: ${error instanceof Error ? error.message : 'Storage error'}`,
			{ cause: error },
		);
	}
	checkFileAbort(options?.signal);
	const attachment = sanitizeAttachment({ ...uploaded, name });
	if (!attachment?.url || !contentUrl(attachment.url, 'https://workspace.invalid'))
		throw new Error('The file could not be uploaded');
	return attachment;
}
