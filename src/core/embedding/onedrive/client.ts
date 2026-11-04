import { OPEN_SELECT, itemUrl, listingFromPage, pagePath, searchPath } from './browse';
import { copyName, fileFromMetadata, validateFile } from './metadata';
import {
	checkResponse,
	encodeOneDriveSharingUrl,
	readBytes,
	readJson,
	secureTransferUrl,
} from './requests';
import { createTransport } from './transport';
import { uploadBytes } from './upload';
import { OneDriveError } from './types';
import type {
	OneDriveClient,
	OneDriveClientOptions,
	OneDriveFile,
	OneDriveOpenOptions,
	OneDriveOpened,
} from './types';

const itemPath = (file: OneDriveFile) => itemUrl(file);
const openPath = (file: { driveId: string; itemId: string }) =>
	`${itemUrl(file)}?$select=${OPEN_SELECT}`;
/** A DOM-free host-owned Graph connection for every Office editor. */
export function createOneDriveClient(options: OneDriveClientOptions): OneDriveClient {
	const transport = createTransport(options);
	const { operation, graph } = transport;
	async function access(file: OneDriveFile) {
		if (options.canWrite)
			file.access = (await options.canWrite(Object.freeze({ ...file }))) ? 'write' : 'read';
		return file;
	}
	async function writable(file: OneDriveFile) {
		validateFile(file);
		if (file.access !== 'write' || !(await options.canWrite?.(Object.freeze({ ...file }))))
			throw new OneDriveError(
				'access-denied',
				'This file is read-only. Keep or download your edits instead.',
			);
	}
	async function download(
		data: Record<string, unknown>,
		signal: AbortSignal,
		onProgress?: OneDriveOpenOptions['onProgress'],
	): Promise<OneDriveOpened> {
		const file = fileFromMetadata(data);
		if (typeof data.size === 'number' && data.size > transport.maxBytes)
			throw new OneDriveError('too-large', 'The file exceeds the configured size limit.');
		const url = secureTransferUrl(data['@microsoft.graph.downloadUrl']);
		// The host's write policy resolves while the bytes stream; neither waits for the other.
		const [, bytes] = await Promise.all([
			access(file),
			transport
				.fetcher(url, { signal, credentials: 'omit', referrerPolicy: 'no-referrer' })
				.then((response) => readBytes(checkResponse(response), transport.maxBytes, onProgress)),
		]);
		if (!bytes.length) throw new OneDriveError('unsupported', 'The Office file is empty.');
		return { file, bytes };
	}
	const page = (path: string, input: Parameters<OneDriveClient['recent']>[0]) =>
		operation(input, async (signal) =>
			listingFromPage(await readJson(await graph(pagePath(path, input), signal))),
		);
	return {
		open(shareUrl, input) {
			const encoded = encodeOneDriveSharingUrl(shareUrl);
			return operation(input, async (signal) =>
				download(
					await readJson(
						await graph(`/shares/${encoded}/driveItem?$select=${OPEN_SELECT}`, signal, {
							// Lets a Files.Read token open a link without a separate redemption call.
							headers: { Prefer: 'redeemSharingLink' },
						}),
					),
					signal,
					input?.onProgress,
				),
			);
		},
		openItem(item, input) {
			const target = { driveId: item.driveId, itemId: item.itemId };
			return operation(input, async (signal) => {
				if (!target.driveId || !target.itemId)
					throw new OneDriveError('unsupported', 'Choose a OneDrive file.');
				const opened = await download(
					await readJson(await graph(openPath(target), signal)),
					signal,
					input?.onProgress,
				);
				if (opened.file.driveId !== target.driveId || opened.file.itemId !== target.itemId)
					throw new OneDriveError('invalid-response', 'OneDrive returned a different file.');
				return opened;
			});
		},
		metadata(file, input) {
			const target = { ...file };
			return operation(input, async (signal) => {
				validateFile(target);
				const data = await readJson(
					await graph(
						`${itemPath(target)}?$select=id,name,eTag,webUrl,parentReference,file`,
						signal,
					),
				);
				const current = await access(fileFromMetadata(data, target));
				if (current.itemId !== target.itemId || current.driveId !== target.driveId)
					throw new OneDriveError('invalid-response', 'OneDrive returned a different file.');
				return current;
			});
		},
		list(folder, input) {
			const path = folder ? `${itemUrl(folder)}/children` : '/me/drive/root/children';
			return page(path, input);
		},
		recent: (input) => page('/me/drive/recent', input),
		search(query, input) {
			if (!query.trim()) return Promise.resolve({ entries: [] });
			return page(searchPath(query), input);
		},
		read(file, input) {
			const target = { ...file };
			return operation(input, async (signal) => {
				validateFile(target);
				const opened = await download(
					await readJson(await graph(openPath(target), signal)),
					signal,
					input?.onProgress,
				);
				if (
					opened.file.driveId !== target.driveId ||
					opened.file.itemId !== target.itemId ||
					opened.file.kind !== target.kind
				)
					throw new OneDriveError('invalid-response', 'OneDrive returned a different file.');
				return opened;
			});
		},
		save(file, bytes, input) {
			const target = { ...file };
			const snapshot = bytes.slice();
			return operation(input, async (signal) => {
				await writable(target);
				const updated = fileFromMetadata(
					await uploadBytes(transport, itemPath(target), snapshot, signal, target.eTag),
					target,
				);
				if (
					updated.itemId !== target.itemId ||
					updated.driveId !== target.driveId ||
					updated.kind !== target.kind
				)
					throw new OneDriveError(
						'invalid-response',
						'OneDrive returned a different file after saving.',
					);
				return access(updated);
			});
		},
		saveCopy(file, bytes, name, input) {
			const target = { ...file };
			const snapshot = bytes.slice();
			const candidate = copyName(target, name);
			return operation(input, async (signal) => {
				await writable(target);
				if (!target.parentId)
					throw new OneDriveError(
						'unsupported',
						'Reopen this file to obtain its folder before saving a cloud copy.',
					);
				const path = `/drives/${encodeURIComponent(target.driveId)}/items/${encodeURIComponent(target.parentId)}:/${encodeURIComponent(candidate)}:`;
				const updated = fileFromMetadata(await uploadBytes(transport, path, snapshot, signal));
				if (
					updated.itemId === target.itemId ||
					updated.driveId !== target.driveId ||
					updated.kind !== target.kind
				)
					throw new OneDriveError('invalid-response', 'OneDrive did not create a separate copy.');
				return { file: await access(updated), bytes: snapshot };
			});
		},
		share(file, link, input) {
			const target = { ...file };
			const requested = { ...link };
			return operation(input, async (signal) => {
				await writable(target);
				if (
					!['view', 'edit'].includes(requested.type) ||
					!['anonymous', 'organization', 'users'].includes(requested.scope)
				)
					throw new OneDriveError(
						'unsupported',
						'Choose an explicit sharing permission and audience.',
					);
				const data = await readJson(
					await graph(itemPath(target) + '/createLink', signal, {
						method: 'POST',
						headers: { 'Content-Type': 'application/json' },
						body: JSON.stringify(requested),
					}),
				);
				const result = data.link as Record<string, unknown> | undefined;
				return secureTransferUrl(result?.webUrl);
			});
		},
	};
}
