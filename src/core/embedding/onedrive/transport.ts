import { checkResponse } from './requests';
import { OneDriveError } from './types';
import type { OneDriveClientOptions, OneDriveRequestOptions } from './types';

const GRAPH = 'https://graph.microsoft.com/v1.0';
/** Only Graph itself ever receives the bearer token, including for paging links. */
export function graphUrl(value: unknown): string {
	try {
		const url = new URL(String(value));
		if (
			url.protocol === 'https:' &&
			url.hostname === 'graph.microsoft.com' &&
			!url.username &&
			!url.password &&
			!url.port
		)
			return url.href;
	} catch {
		// Fall through without echoing the link.
	}
	throw new OneDriveError('invalid-response', 'OneDrive returned an invalid paging link.');
}
/** Graph requires upload fragments in multiples of 320 KiB, each request below 60 MiB. */
const UPLOAD_FRAGMENT_UNIT = 320 * 1024;
const UPLOAD_FRAGMENT_MAX = 191 * UPLOAD_FRAGMENT_UNIT;
export function createTransport(options: OneDriveClientOptions) {
	// Transfer calls go through the transport object; browser fetch requires its Window receiver.
	const fetcher = options.fetch ?? globalThis.fetch.bind(globalThis);
	const maxBytes = options.maxBytes ?? 32 * 1024 * 1024;
	const timeoutMs = options.timeoutMs ?? 60000;
	const fragmentBytes = options.uploadFragmentBytes ?? UPLOAD_FRAGMENT_MAX;
	if (
		!Number.isSafeInteger(maxBytes) ||
		maxBytes <= 0 ||
		!Number.isFinite(timeoutMs) ||
		timeoutMs <= 0
	)
		throw new RangeError('OneDrive limits must be positive.');
	if (
		!Number.isSafeInteger(fragmentBytes) ||
		fragmentBytes <= 0 ||
		fragmentBytes % UPLOAD_FRAGMENT_UNIT !== 0 ||
		fragmentBytes > UPLOAD_FRAGMENT_MAX
	)
		throw new RangeError('OneDrive upload fragments must be a multiple of 320 KiB below 60 MiB.');
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
	// One token request per operation: parallel and sequential Graph calls of the same operation
	// share it. Keyed by the operation's signal, so nothing outlives the operation or is stored.
	const tokens = new WeakMap<AbortSignal, Promise<string>>();
	function token(signal: AbortSignal): Promise<string> {
		let pending = tokens.get(signal);
		if (!pending) {
			pending = Promise.resolve()
				.then(() => options.getAccessToken())
				.then(
					(value) => {
						if (!value) throw new Error('empty');
						return value;
					},
					() => {
						throw new Error('failed');
					},
				);
			tokens.set(signal, pending);
		}
		return pending;
	}
	/** `path` is a v1.0 path, or an absolute Graph URL (a validated `@odata.nextLink`). */
	async function graph(path: string, signal: AbortSignal, init: RequestInit = {}) {
		const url = path.startsWith('/') ? GRAPH + path : graphUrl(path);
		let token_: string;
		try {
			token_ = await token(signal);
		} catch {
			throw new OneDriveError('authentication', 'Sign in to Microsoft to access this file.');
		}
		signal.throwIfAborted();
		const headers = new Headers(init.headers);
		headers.set('Authorization', `Bearer ${token_}`);
		return checkResponse(
			await fetcher(url, {
				...init,
				headers,
				signal,
				credentials: 'omit',
				redirect: 'error',
				referrerPolicy: 'no-referrer',
			}),
		);
	}
	return { operation, graph, maxBytes, fetcher, fragmentBytes };
}
export type OneDriveTransport = ReturnType<typeof createTransport>;
