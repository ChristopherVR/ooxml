import { checkResponse, readJson, secureTransferUrl } from './requests';
import { OneDriveError } from './types';
import type { OneDriveTransport } from './transport';

export async function uploadBytes(
	transport: OneDriveTransport,
	path: string,
	bytes: Uint8Array,
	signal: AbortSignal,
	eTag?: string,
	onProgress?: (sent: number, total: number) => void,
): Promise<Record<string, unknown>> {
	if (!bytes.length || bytes.length > transport.maxBytes)
		throw new OneDriveError(
			'too-large',
			'Choose nonempty Office file bytes within the configured size limit.',
		);
	const session = await readJson(
		await transport.graph(path + '/createUploadSession', signal, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json', ...(eTag ? { 'If-Match': eTag } : {}) },
			body: JSON.stringify({
				item: { '@microsoft.graph.conflictBehavior': eTag ? 'replace' : 'rename' },
			}),
		}),
	);
	const uploadUrl = secureTransferUrl(session.uploadUrl);
	try {
		return await sendFragments(transport, uploadUrl, bytes, signal, onProgress);
	} catch (error) {
		// An abandoned session blocks the next one for this item (nameAlreadyExists), so cancel it.
		await transport
			.fetcher(uploadUrl, { method: 'DELETE', credentials: 'omit', referrerPolicy: 'no-referrer' })
			.catch(() => undefined);
		throw error;
	}
}

/** Graph reports the next range as `start-` or `start-end`; only the start matters. */
function nextRangeStart(value: unknown): number | undefined {
	const match = typeof value === 'string' ? /^(\d+)-/u.exec(value) : null;
	return match ? Number(match[1]) : undefined;
}

async function sendFragments(
	transport: OneDriveTransport,
	uploadUrl: string,
	bytes: Uint8Array,
	signal: AbortSignal,
	onProgress?: (sent: number, total: number) => void,
): Promise<Record<string, unknown>> {
	onProgress?.(0, bytes.length);
	// One request per fragment (`uploadFragmentBytes`, 320 KiB aligned; a single one by default).
	for (let start = 0; start < bytes.length; start += transport.fragmentBytes) {
		const end = Math.min(start + transport.fragmentBytes, bytes.length);
		const response = checkResponse(
			await transport.fetcher(uploadUrl, {
				method: 'PUT',
				signal,
				credentials: 'omit',
				referrerPolicy: 'no-referrer',
				headers: { 'Content-Range': `bytes ${start}-${end - 1}/${bytes.length}` },
				body: new Uint8Array(bytes.subarray(start, end)),
			}),
		);
		const data = await readJson(response);
		onProgress?.(end, bytes.length);
		if (end === bytes.length) {
			if (response.status === 200 || response.status === 201) return data;
			throw new OneDriveError('invalid-response', 'OneDrive did not confirm the completed upload.');
		}
		if (
			response.status !== 202 ||
			!Array.isArray(data.nextExpectedRanges) ||
			nextRangeStart(data.nextExpectedRanges[0]) !== end
		)
			throw new OneDriveError('invalid-response', 'OneDrive requested an unexpected upload range.');
	}
	throw new OneDriveError('invalid-response', 'OneDrive did not complete the upload.');
}
