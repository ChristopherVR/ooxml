import { checkResponse, readJson, secureTransferUrl } from './requests';
import { OneDriveError } from './types';
import type { OneDriveTransport } from './transport';

// Every non-final fragment is a multiple of Graph's required 320 KiB.
const CHUNK_SIZE = 5 * 1024 * 1024;
export async function uploadBytes(
	transport: OneDriveTransport,
	path: string,
	bytes: Uint8Array,
	signal: AbortSignal,
	eTag?: string,
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
				item: { '@microsoft.graph.conflictBehavior': eTag ? 'fail' : 'rename' },
			}),
		}),
	);
	const uploadUrl = secureTransferUrl(session.uploadUrl);
	for (let start = 0; start < bytes.length; start += CHUNK_SIZE) {
		const end = Math.min(start + CHUNK_SIZE, bytes.length);
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
		if (end === bytes.length) {
			if (response.status === 200 || response.status === 201) return data;
			throw new OneDriveError('invalid-response', 'OneDrive did not confirm the completed upload.');
		}
		if (
			response.status !== 202 ||
			!Array.isArray(data.nextExpectedRanges) ||
			data.nextExpectedRanges[0] !== `${end}-`
		)
			throw new OneDriveError('invalid-response', 'OneDrive requested an unexpected upload range.');
	}
	throw new OneDriveError('invalid-response', 'OneDrive did not complete the upload.');
}
