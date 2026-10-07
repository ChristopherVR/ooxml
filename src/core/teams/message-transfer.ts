import { createClientId } from '../collab/identity';
import { MAX_ATTACHMENTS, type Attachment } from './model';
import type { DraftContext } from './drafts';
import { checkFileAbort, type FileTransferProgress } from './file-transfer';
import { storeAttachment, storageFileName, validateStorageFile } from './stored-file';
import type { UploadableFile } from './store';

export interface MessageTransfer {
	id: string;
	context: DraftContext;
	progress: FileTransferProgress;
}
export function createMessageTransfers(context: {
	available(channelId: string): boolean;
	canUpload(): boolean;
	upload(file: UploadableFile & Blob, signal?: AbortSignal): Promise<Attachment>;
	changed(): void;
}) {
	const active = new Map<string, { view: MessageTransfer; controller: AbortController }>();
	return {
		list: (): MessageTransfer[] =>
			[...active.values()].map(({ view }) => ({
				...view,
				context: { ...view.context },
				progress: { ...view.progress },
			})),
		cancel(id: string): void {
			active.get(id)?.controller.abort();
		},
		destroy(): void {
			for (const item of active.values()) item.controller.abort();
		},
		async upload(
			destination: DraftContext,
			files: (UploadableFile & Blob)[],
			publish: (attachments: Attachment[]) => void,
		): Promise<void> {
			if (!context.available(destination.channelId))
				throw new Error('The channel is no longer available');
			if (!context.canUpload())
				throw new Error('Configure file storage before sharing attachments');
			if (!files.length || files.length > MAX_ATTACHMENTS)
				throw new Error(`Choose between 1 and ${MAX_ATTACHMENTS} files`);
			files.forEach(validateStorageFile);
			const id = createClientId();
			const controller = new AbortController();
			const view: MessageTransfer = {
				id,
				context: { ...destination },
				progress: {
					phase: 'uploading',
					fileName: storageFileName(files[0]!.name),
					completed: 0,
					total: files.length,
				},
			};
			active.set(id, { view, controller });
			context.changed();
			const attachments: Attachment[] = [];
			try {
				for (const file of files) {
					checkFileAbort(controller.signal);
					if (!context.available(destination.channelId))
						throw new Error('The channel is no longer available');
					view.progress = {
						phase: 'uploading',
						fileName: storageFileName(file.name),
						completed: attachments.length,
						total: files.length,
					};
					context.changed();
					attachments.push(
						await storeAttachment(file, 'message', context.upload, { signal: controller.signal }),
					);
					view.progress = { ...view.progress, phase: 'uploaded', completed: attachments.length };
					context.changed();
				}
				checkFileAbort(controller.signal);
				if (!context.available(destination.channelId))
					throw new Error('The channel is no longer available; the uploaded files were not shared');
				publish(attachments);
			} finally {
				active.delete(id);
				context.changed();
			}
		},
	};
}
