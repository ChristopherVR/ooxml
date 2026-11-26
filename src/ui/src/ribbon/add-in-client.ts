import { ADD_IN_PROTOCOL, ADD_IN_TEXT_LIMIT, ADD_IN_TIMEOUT, HOST_NAMES } from './add-in-protocol';
import type { OfficeHostName } from './add-in-protocol';

export type OfficeAddInResult<T> = { asyncContext: unknown } & (
	| { status: 'succeeded'; value: T }
	| { status: 'failed'; error: { code: string; message: string } }
);
type Callback<T> = (result: OfficeAddInResult<T>) => void;
type CallOptions = { asyncContext?: unknown; coercionType?: string };
type Ready = { host: OfficeHostName; platform: 'OfficeOnline' };

const unreachable = 'The editor did not respond to the add-in.';

function unsupportedMembers<T extends object>(surface: T, path: string): T {
	return new Proxy(surface, {
		get(target, property, receiver) {
			if (typeof property === 'string' && property !== 'then' && !Object.hasOwn(target, property))
				throw new Error(`Unsupported Office.js member: ${path}.${property}`);
			return Reflect.get(target, property, receiver);
		},
	});
}

/** Connect the compatible client from inside the publisher's sandboxed frame. */
export function connectOfficeAddIn(parentOrigin: string) {
	const target = new URL(parentOrigin);
	if (!['https:', 'http:'].includes(target.protocol) || target.origin !== parentOrigin)
		throw new TypeError('An explicit editor origin is required.');
	if (window.parent === window) throw new Error('This add-in needs an editor frame.');
	const { port1, port2 } = new MessageChannel();
	const waiting = new Map<number, (result: { value: string | null } | { error: string }) => void>();
	let nextId = 0;
	let info: Ready | undefined;
	const ready = new Promise<Ready>((resolve, reject) => {
		const timer = window.setTimeout(() => {
			port1.close();
			reject(new Error(unreachable));
		}, ADD_IN_TIMEOUT);
		port1.onmessage = (event: MessageEvent<unknown>) => {
			const data = event.data;
			if (!data || typeof data !== 'object') return;
			if (
				!info &&
				'kind' in data &&
				data.kind === 'ready' &&
				'host' in data &&
				typeof data.host === 'string'
			) {
				const host = Object.values(HOST_NAMES).find((name) => name === data.host);
				if (!host) return;
				window.clearTimeout(timer);
				info = { host, platform: 'OfficeOnline' };
				resolve(info);
			}
			if (!info || !('id' in data) || typeof data.id !== 'number') return;
			const complete = waiting.get(data.id);
			if ('error' in data && typeof data.error === 'string') complete?.({ error: data.error });
			else if ('value' in data && (typeof data.value === 'string' || data.value === null))
				complete?.({ value: data.value });
		};
		window.parent.postMessage({ protocol: ADD_IN_PROTOCOL, kind: 'connect' }, parentOrigin, [
			port2,
		]);
	});
	// Consumers may attach an onReady handler after the initialization timeout.
	void ready.catch(() => {});

	async function call<T>(
		method: string,
		value: string | undefined,
		options: CallOptions | null | Callback<T>,
		callback: Callback<T> | undefined,
		localError: string | undefined,
		convert: (value: string | null) => T,
	) {
		const handler = typeof options === 'function' ? options : callback;
		const asyncContext = typeof options === 'function' ? undefined : options?.asyncContext;
		const failure = (message: string): OfficeAddInResult<T> => ({
			status: 'failed',
			asyncContext,
			error: { code: 'OfficeAddInError', message },
		});
		let result: OfficeAddInResult<T>;
		try {
			if (localError) throw new Error(localError);
			await ready;
			if (waiting.size >= 16) throw new Error('Too many Office.js calls are pending.');
			const response = await new Promise<{ value: string | null } | { error: string }>(
				(resolve) => {
					const id = ++nextId;
					const timer = window.setTimeout(() => {
						waiting.delete(id);
						resolve({ error: unreachable });
					}, ADD_IN_TIMEOUT);
					waiting.set(id, (result) => {
						waiting.delete(id);
						window.clearTimeout(timer);
						resolve(result);
					});
					port1.postMessage({ id, method, ...(value !== undefined && { value }) });
				},
			);
			result =
				'error' in response
					? failure(response.error)
					: { status: 'succeeded', asyncContext, value: convert(response.value) };
		} catch (error) {
			result = failure(error instanceof Error ? error.message : unreachable);
		}
		if (handler) handler(result);
		else if (result.status === 'failed') throw new Error(result.error.message);
	}

	const selection = unsupportedMembers(
		{
			getSelectedDataAsync(
				coercionType: string,
				options: CallOptions | null | Callback<string>,
				callback?: Callback<string>,
			) {
				return call(
					'getSelectedDataAsync',
					undefined,
					options,
					callback,
					coercionType === 'text' ? undefined : 'Only text coercion is supported.',
					(value) => value ?? '',
				);
			},
			setSelectedDataAsync(
				text: string,
				options: CallOptions | null | Callback<undefined>,
				callback?: Callback<undefined>,
			) {
				return call(
					'setSelectedDataAsync',
					text,
					options,
					callback,
					typeof text !== 'string' ||
						text.length > ADD_IN_TEXT_LIMIT ||
						(typeof options !== 'function' &&
							options?.coercionType &&
							options.coercionType !== 'text')
						? 'Only bounded plain text is supported.'
						: undefined,
					() => undefined,
				);
			},
		},
		'Office.context.document',
	);

	return unsupportedMembers(
		{
			onReady: (callback?: (info: Ready) => void) =>
				ready.then((info) => {
					callback?.(info);
					return info;
				}),
			context: unsupportedMembers(
				{
					get host() {
						return info?.host;
					},
					get platform() {
						return info?.platform;
					},
					document: selection,
					requirements: { isSetSupported: () => false },
				},
				'Office.context',
			),
			CoercionType: { Text: 'text' },
			AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
			HostType: { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint' },
			PlatformType: { OfficeOnline: 'OfficeOnline' },
		},
		'Office',
	);
}
