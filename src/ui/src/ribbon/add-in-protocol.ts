import type { OfficeAddInHost } from 'ooxml-core/automation/add-ins';
import type { OfficeAddInSelection } from './add-in-runtime-settings';

export const ADD_IN_PROTOCOL = 'ooxml-office-add-in:1';
export const ADD_IN_TEXT_LIMIT = 32_768;
export const ADD_IN_TIMEOUT = 10_000;
export const HOST_NAMES = { word: 'Word', excel: 'Excel', powerpoint: 'PowerPoint' } as const;
export type OfficeHostName = (typeof HOST_NAMES)[OfficeAddInHost];

const READ = new Set(['Document.Read.User', 'Document.ReadAll.User', 'Document.ReadWrite.User']);
const WRITE = new Set(['Document.Write.User', 'Document.ReadWrite.User']);

export interface BridgeRequest {
	id: number;
	method: string;
	value?: string;
}

export type BridgeResult = { value: string | null } | { error: string };

export function bridgeRequest(value: unknown): BridgeRequest | undefined {
	if (
		!value ||
		typeof value !== 'object' ||
		!('id' in value) ||
		typeof value.id !== 'number' ||
		!Number.isSafeInteger(value.id) ||
		value.id < 1 ||
		!('method' in value) ||
		typeof value.method !== 'string' ||
		value.method.length > 64
	)
		return;
	const id = value.id;
	const method = value.method;
	if (method === 'setSelectedDataAsync') {
		if (
			!('value' in value) ||
			typeof value.value !== 'string' ||
			value.value.length > ADD_IN_TEXT_LIMIT
		)
			return;
		return { id, method, value: value.value };
	}
	if (!('value' in value)) return { id, method };
}

export function dispatchAddInRequest(
	request: BridgeRequest,
	permissions: readonly string[],
	selection: OfficeAddInSelection | undefined,
): BridgeResult {
	switch (request.method) {
		case 'getSelectedDataAsync': {
			if (!permissions.some((permission) => READ.has(permission)))
				return { error: 'This add-in has no permission to read the selection.' };
			const value = selection?.readText();
			return value !== undefined && value.length <= ADD_IN_TEXT_LIMIT
				? { value }
				: { error: 'Text selection is unavailable or exceeds the size limit.' };
		}
		case 'setSelectedDataAsync':
			if (!permissions.some((permission) => WRITE.has(permission)))
				return { error: 'This add-in has no permission to change the selection.' };
			return request.value !== undefined && selection?.writeText(request.value)
				? { value: null }
				: { error: 'The current selection cannot be edited as text.' };
		default:
			return { error: 'This Office.js operation is unsupported.' };
	}
}

export function isPaneHandshake(event: MessageEvent, source: Window | null): boolean {
	const data: unknown = event.data;
	return (
		!!source &&
		event.source === source &&
		event.origin === 'null' &&
		event.ports.length === 1 &&
		!!data &&
		typeof data === 'object' &&
		'protocol' in data &&
		data.protocol === ADD_IN_PROTOCOL &&
		'kind' in data &&
		data.kind === 'connect'
	);
}
