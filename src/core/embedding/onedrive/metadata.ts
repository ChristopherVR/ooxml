import { OneDriveError } from './types';
import type { OneDriveFile, OneDriveKind } from './types';

export function officeKind(name: string): OneDriveKind {
	const extension = /\.(docx|xlsx|pptx|vsdx)$/iu.exec(name)?.[1]?.toLowerCase();
	if (!extension)
		throw new OneDriveError('unsupported', 'Choose a .docx, .xlsx, .pptx or .vsdx file.');
	return extension as OneDriveKind;
}

export function fileFromMetadata(
	data: Record<string, unknown>,
	previous?: OneDriveFile,
): OneDriveFile {
	const parent = data.parentReference as Record<string, unknown> | undefined;
	const driveId = parent?.driveId ?? previous?.driveId;
	const parentId = parent?.id ?? previous?.parentId;
	const { id: itemId, name, eTag } = data;
	const webUrl = data.webUrl ?? previous?.webUrl;
	if (
		![driveId, itemId, name, eTag].every((value) => typeof value === 'string' && value.length > 0)
	)
		throw new OneDriveError(
			'invalid-response',
			'OneDrive did not return a complete file identity.',
		);
	if (data.folder || data.package || data.remoteItem)
		throw new OneDriveError(
			'unsupported',
			'Use a sharing link to an Office file, rather than a folder or shortcut.',
		);
	let url: URL | undefined;
	if (webUrl !== undefined) {
		try {
			url = new URL(String(webUrl));
		} catch {
			throw new OneDriveError('invalid-response', 'OneDrive returned an invalid file address.');
		}
		if (url.protocol !== 'https:' || url.username || url.password)
			throw new OneDriveError('invalid-response', 'OneDrive returned an invalid file address.');
	}
	return {
		driveId: driveId as string,
		itemId: itemId as string,
		name: name as string,
		kind: officeKind(name as string),
		eTag: eTag as string,
		access: 'unknown',
		...(url ? { webUrl: url.href } : {}),
		...(typeof parentId === 'string' && parentId ? { parentId } : {}),
	};
}

export function validateFile(file: OneDriveFile): void {
	const checked = fileFromMetadata({
		id: file.itemId,
		name: file.name,
		eTag: file.eTag,
		webUrl: file.webUrl,
		parentReference: { driveId: file.driveId },
	});
	if (file.kind !== checked.kind)
		throw new OneDriveError('unsupported', 'The file format does not match its identity.');
}

export function copyName(file: OneDriveFile, name?: string): string {
	const candidate = name ?? file.name.replace(/(\.[^.]+)$/u, ' (copy)$1');
	if (
		!candidate ||
		/[\\/:*?"<>|\u0000-\u001f]/u.test(candidate) ||
		candidate.endsWith('.') ||
		candidate.endsWith(' ') ||
		candidate.length > 255 ||
		officeKind(candidate) !== file.kind
	)
		throw new OneDriveError(
			'unsupported',
			'Use a valid copy filename with the same Office extension.',
		);
	return candidate;
}
