import { checkResponse } from './requests';
import { OneDriveError } from './types';
import type { OneDriveClientOptions, OneDriveRequestOptions } from './types';

const GRAPH = 'https://graph.microsoft.com/v1.0';
export function createTransport(options: OneDriveClientOptions) {
	// Transfer calls go through the transport object; browser fetch requires its Window receiver.
	const fetcher = options.fetch ?? globalThis.fetch.bind(globalThis);
	const maxBytes = options.maxBytes ?? 32 * 1024 * 1024;
	const timeoutMs = options.timeoutMs ?? 60000;
	if (
		!Number.isSafeInteger(maxBytes) ||
		maxBytes <= 0 ||
		!Number.isFinite(timeoutMs) ||
		timeoutMs <= 0
	)
		throw new RangeError('OneDrive limits must be positive.');
	async function operation<T>(
		input: OneDriveRequestOptions | undefined,
		run: (signal: AbortSignal) => Promise<T>,
	): Promise<T> {
		const signal = AbortSignal.any([
			AbortSignal.timeout(timeoutMs),
			...(input?.signal ? [input.signal] : []),
		]);
		signal.throwIfAborted();
		let abort: () => void = () => {};
		try {
			return await Promise.race([
				run(signal),
				new Promise<never>((_, reject) => {
					abort = () => reject(signal.reason);
					signal.addEventListener('abort', abort, { once: true });
					if (signal.aborted) abort();
				}),
			]);
		} catch (error) {
			if (signal.aborted) throw signal.reason;
			if (error instanceof OneDriveError) throw error;
			throw new OneDriveError(
				'network',
				'Could not connect to OneDrive. Your edits are still available.',
			);
		} finally {
			signal.removeEventListener('abort', abort);
		}
	}
	async function graph(path: string, signal: AbortSignal, init: RequestInit = {}) {
		let token: string;
		try {
			token = await options.getAccessToken();
		} catch {
			throw new OneDriveError('authentication', 'Sign in to Microsoft to access this file.');
		}
		signal.throwIfAborted();
		if (!token)
			throw new OneDriveError('authentication', 'Sign in to Microsoft to access this file.');
		const headers = new Headers(init.headers);
		headers.set('Authorization', `Bearer ${token}`);
		return checkResponse(
			await fetcher(GRAPH + path, {
				...init,
				headers,
				signal,
				credentials: 'omit',
				redirect: 'error',
				referrerPolicy: 'no-referrer',
			}),
		);
	}
	return { operation, graph, maxBytes, fetcher };
}
export type OneDriveTransport = ReturnType<typeof createTransport>;
