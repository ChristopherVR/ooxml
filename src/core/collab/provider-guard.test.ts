import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import { describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import { adaptYjsProvider } from './external-provider';
import { guardProviderUpdates } from './provider-guard';
import { UPDATE_REFUSAL_MESSAGES, UpdateRefusedError } from './refusal';

/** A sync update frame without the outer message type, as y-websocket's handler receives it. */
function syncUpdate(update: Uint8Array): decoding.Decoder {
	const e = encoding.createEncoder();
	encoding.writeVarUint(e, 2);
	encoding.writeVarUint8Array(e, update);
	return decoding.createDecoder(encoding.toUint8Array(e));
}

function manyStructs(count: number): Uint8Array {
	const e = encoding.createEncoder();
	for (const n of [1, count, 7, 0]) encoding.writeVarUint(e, n);
	for (let i = 0; i < count; i++) {
		encoding.writeUint8(e, 0);
		encoding.writeVarUint(e, 1);
	}
	encoding.writeVarUint(e, 0);
	return encoding.toUint8Array(e);
}

describe('guardProviderUpdates', () => {
	it('passes ordinary updates to the provider and drops one over budget', () => {
		const sync = vi.fn();
		const onRejected = vi.fn();
		const provider = { messageHandlers: [sync] };
		expect(guardProviderUpdates(provider, { maxStructs: 100, onRejected })).toBe(true);
		const doc = new Y.Doc();
		doc.getText('t').insert(0, 'hello');
		provider.messageHandlers[0]!({}, syncUpdate(Y.encodeStateAsUpdate(doc)));
		expect(sync).toHaveBeenCalledOnce();
		// The provider still reads the message from where it started.
		expect(decoding.readVarUint(sync.mock.calls[0]![1] as decoding.Decoder)).toBe(2);

		provider.messageHandlers[0]!({}, syncUpdate(manyStructs(1000)));
		expect(sync).toHaveBeenCalledOnce();
		const [error] = onRejected.mock.calls[0]!;
		expect(error).toBeInstanceOf(UpdateRefusedError);
		expect(error).toMatchObject({
			code: 'too-large',
			detail: 'update holds more than 100 structs',
			message: UPDATE_REFUSAL_MESSAGES['too-large'],
		});
	});

	it('stops a wrapped provider at the first refused update, with the reason', () => {
		const sync = vi.fn();
		const disconnect = vi.fn();
		const provider = { messageHandlers: [sync], on: () => {}, disconnect, wsconnected: true };
		const wrapped = adaptYjsProvider(provider);
		const errors: Error[] = [];
		wrapped.on('error', (error) => errors.push(error));
		provider.messageHandlers[0]!({}, syncUpdate(manyStructs(100_000)));
		expect(sync).not.toHaveBeenCalled();
		expect(disconnect).toHaveBeenCalledOnce();
		expect(wrapped.status).toBe('error');
		expect(errors).toHaveLength(1);
		expect(errors[0]).toMatchObject({ code: 'too-large' });
	});

	it('leaves a provider without a handler table alone', () => {
		expect(guardProviderUpdates({})).toBe(false);
	});
});
