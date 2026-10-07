import { createClientId } from '../collab/identity.js';
import { contentUrl } from './content.js';
import {
	MAX_ATTACHMENTS,
	sanitizeAttachment,
	sanitizeChannelName,
	type Attachment,
} from './model.js';
import type { TeamsClient, UploadableFile } from './store.js';

type FileBlob = UploadableFile & Blob;
interface FileContext {
	available(channelId: string): boolean;
	canUpload(): boolean;
	upload(file: FileBlob): Promise<Attachment>;
	post(channelId: string, text: string, attachments: Attachment[]): boolean;
}

/** Storage operations always capture the destination and use fresh storage names. */
export function createFileActions(
	context: FileContext,
): Pick<TeamsClient, 'saveFileCopy' | 'uploadFiles' | 'createWorkbook'> {
	function ready(channelId: string): void {
		if (!context.available(channelId)) throw new Error('The channel is no longer available');
		if (!context.canUpload()) throw new Error('Configure file storage before sharing files');
	}
	function fileName(raw: string): string {
		return sanitizeChannelName(raw).replace(/[\/\\\u0000-\u001f\u007f]/gu, '');
	}
	async function stored(file: FileBlob, prefix: string): Promise<Attachment> {
		const name = fileName(file.name);
		if (!name || file.size > 33_554_432)
			throw new Error('This file cannot be saved to the channel');
		const ext = name.includes('.') ? `.${name.split('.').pop()}` : '';
		const unique = Object.assign(new Blob([file], { type: file.type ?? '' }), {
			name: `${prefix}-${createClientId()}${ext}`,
		});
		const uploaded = await context.upload(unique);
		const attachment = sanitizeAttachment({ ...uploaded, name });
		if (!attachment?.url || !contentUrl(attachment.url, 'https://workspace.invalid'))
			throw new Error('The file could not be uploaded');
		return attachment;
	}
	function shared(channelId: string, text: string, attachments: Attachment[]): void {
		if (!context.available(channelId))
			throw new Error('The channel is no longer available; the uploaded files were not shared');
		if (!context.post(channelId, text, attachments))
			throw new Error('The files could not be shared');
	}
	return {
		async saveFileCopy(channelId, file) {
			ready(channelId);
			const attachment = await stored(file, 'copy');
			shared(channelId, `Saved a copy of ${attachment.name}`, [attachment]);
			return attachment;
		},
		async uploadFiles(channelId, files) {
			ready(channelId);
			if (!files.length || files.length > MAX_ATTACHMENTS)
				throw new Error(`Choose between 1 and ${MAX_ATTACHMENTS} files`);
			if (files.some((file) => !fileName(file.name) || file.size > 33_554_432))
				throw new Error('This file cannot be saved to the channel');
			const attachments: Attachment[] = [];
			for (const file of files) {
				ready(channelId);
				attachments.push(await stored(file, 'file'));
			}
			shared(channelId, 'Shared files', attachments);
			return attachments;
		},
		async createWorkbook(channelId, rawName) {
			ready(channelId);
			const stem = fileName(rawName)
				.replace(/\.xlsx$/iu, '')
				.trim()
				.slice(0, 75);
			if (!stem) throw new Error('Enter a workbook name');
			const name = `${stem}.xlsx`;
			// SpreadsheetML creation and serialization remain in the Excel engine.
			const [{ createWorkbook }, { saveXlsx }] = await Promise.all([
				import('../xlsx/workbook.js'),
				import('../xlsx/write/index.js'),
			]);
			const bytes = await saveXlsx(createWorkbook());
			ready(channelId);
			const file = Object.assign(
				new Blob([new Uint8Array(bytes)], {
					type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				}),
				{ name },
			);
			const attachment = await stored(file, 'workbook');
			shared(channelId, `Created ${name}`, [attachment]);
			return attachment;
		},
	};
}
