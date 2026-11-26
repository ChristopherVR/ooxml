import { expect, it, vi } from 'vitest';
import {
	ADD_IN_PROTOCOL,
	ADD_IN_TEXT_LIMIT,
	bridgeRequest,
	dispatchAddInRequest,
	isPaneHandshake,
} from './add-in-protocol';

it('checks both the opaque origin and the exact pane window before accepting a port', () => {
	const channel = new MessageChannel();
	try {
		const source = window;
		const connect = (origin: string, protocol = ADD_IN_PROTOCOL) =>
			new MessageEvent('message', {
				origin,
				source,
				data: { protocol, kind: 'connect' },
				ports: [channel.port1],
			});
		expect(isPaneHandshake(connect('null'), source)).toBe(true);
		expect(isPaneHandshake(connect('https://pane.example.test'), source)).toBe(false);
		expect(isPaneHandshake(connect('null', 'other'), source)).toBe(false);
		expect(isPaneHandshake(connect('null'), null)).toBe(false);
		expect(isPaneHandshake(new MessageEvent('message'), source)).toBe(false);
	} finally {
		channel.port1.close();
		channel.port2.close();
	}
});

it('checks requests before invoking the adapter and applies independent read and write permissions', () => {
	const selection = { readText: vi.fn(() => 'selected text'), writeText: vi.fn(() => true) };
	const read = { id: 1, method: 'getSelectedDataAsync' };
	const write = { id: 2, method: 'setSelectedDataAsync', value: 'changed' };
	expect(bridgeRequest({ ...write, value: 'x'.repeat(ADD_IN_TEXT_LIMIT + 1) })).toBeUndefined();
	expect(bridgeRequest({ ...write, value: { text: 'changed' } })).toBeUndefined();
	expect(bridgeRequest({ ...read, value: '' })).toBeUndefined();
	expect(bridgeRequest({ ...read, id: 0 })).toBeUndefined();
	expect(bridgeRequest(read)).toEqual(read);
	expect(bridgeRequest(write)).toEqual(write);
	expect(dispatchAddInRequest(read, ['Document.Write.User'], selection)).toHaveProperty('error');
	expect(dispatchAddInRequest(write, ['Document.Read.User'], selection)).toHaveProperty('error');
	expect(selection.readText).not.toHaveBeenCalled();
	expect(selection.writeText).not.toHaveBeenCalled();
	expect(dispatchAddInRequest(read, ['Document.Read.User'], selection)).toEqual({
		value: 'selected text',
	});
	expect(dispatchAddInRequest(write, ['Document.Write.User'], selection)).toEqual({ value: null });
	expect(selection.writeText).toHaveBeenCalledWith('changed');
	expect(dispatchAddInRequest({ id: 3, method: 'getAccessToken' }, [], selection)).toEqual({
		error: 'This Office.js operation is unsupported.',
	});
});
