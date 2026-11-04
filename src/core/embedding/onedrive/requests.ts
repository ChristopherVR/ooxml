import { OneDriveError } from './types';

/** Share links are capabilities: do not log them or include them in error text. */
export function encodeOneDriveSharingUrl(raw: string): string {
	let url: URL;
	try {
		url = new URL(raw.trim());
	} catch {
		throw new OneDriveError('invalid-link', 'Enter a valid OneDrive or SharePoint sharing link.');
	}
	const host = url.hostname.toLowerCase();
	const allowed =
		host === '1drv.ms' || host === 'onedrive.live.com' || host.endsWith('.sharepoint.com');
	if (!allowed || url.protocol !== 'https:' || url.username || url.password || url.port) {
		throw new OneDriveError('invalid-link', 'Use an HTTPS OneDrive or SharePoint sharing link.');
	}
	url.hash = '';
	const bytes = new TextEncoder().encode(url.href);
	let binary = '';
	for (const byte of bytes) binary += String.fromCharCode(byte);
	return 'u!' + btoa(binary).replace(/=/gu, '').replace(/\+/gu, '-').replace(/\//gu, '_');
}

export function secureTransferUrl(value: unknown): string {
	try {
		const url = new URL(String(value));
		if (url.protocol === 'https:' && !url.username && !url.password) return url.href;
	} catch {
		// Do not expose the preauthenticated URL in the error.
	}
	throw new OneDriveError('invalid-response', 'OneDrive returned an invalid transfer URL.');
}

export function checkResponse(response: Response): Response {
	if (response.ok) return response;
	const status = response.status;
	if (status === 401)
		throw new OneDriveError('authentication', 'Sign in to Microsoft again.', status);
	if (status === 403)
		throw new OneDriveError('access-denied', 'Your account cannot access this file.', status);
	if (status === 409 || status === 412)
		throw new OneDriveError(
			'conflict',
			'The OneDrive file changed. Reopen it before saving.',
			status,
		);
	if (status === 423)
		throw new OneDriveError(
			'locked',
			'The OneDrive file is locked. Keep your edits and try again later.',
			status,
		);
	throw new OneDriveError('network', 'The OneDrive request failed. Try again.', status);
}

/**
 * Bound streamed bodies even if the server omits or misreports Content-Length. The buffer is
 * preallocated from a trusted-looking Content-Length and grows geometrically otherwise, so each
 * byte is copied once (no chunk list and no final concatenation pass).
 */
export async function readBytes(
	response: Response,
	limit: number,
	onProgress?: (loaded: number, total?: number) => void,
): Promise<Uint8Array> {
	const declared = Number(response.headers.get('Content-Length'));
	if (declared > limit) {
		await response.body?.cancel();
		throw new OneDriveError('too-large', 'The file exceeds the configured size limit.');
	}
	const reader = response.body?.getReader();
	if (!reader) throw new OneDriveError('invalid-response', 'OneDrive returned an empty response.');
	const total = Number.isFinite(declared) && declared > 0 ? declared : undefined;
	let buffer = new Uint8Array(total ?? 64 * 1024);
	let size = 0;
	try {
		while (true) {
			const { value, done } = await reader.read();
			if (done) break;
			size += value.byteLength;
			if (size > limit) {
				await reader.cancel();
				throw new OneDriveError('too-large', 'The file exceeds the configured size limit.');
			}
			if (size > buffer.length) {
				const grown = new Uint8Array(Math.min(Math.max(size, buffer.length * 2), limit));
				grown.set(buffer.subarray(0, size - value.byteLength));
				buffer = grown;
			}
			buffer.set(value, size - value.byteLength);
			onProgress?.(size, total);
		}
	} finally {
		reader.releaseLock();
	}
	return size === buffer.length ? buffer : buffer.slice(0, size);
}

export async function readJson(response: Response): Promise<Record<string, unknown>> {
	try {
		const data: unknown = JSON.parse(
			new TextDecoder().decode(await readBytes(response, 1024 * 1024)),
		);
		if (data && typeof data === 'object' && !Array.isArray(data))
			return data as Record<string, unknown>;
	} catch (error) {
		if (error instanceof OneDriveError) throw error;
	}
	throw new OneDriveError('invalid-response', 'OneDrive returned invalid file metadata.');
}
