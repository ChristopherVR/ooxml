// The update scan for providers that apply sync messages themselves. y-websocket decodes and
// applies every sync message it receives inside its own handler, so an update never passes through
// `applyUpdateSafe`. Its providers keep a per-instance `messageHandlers` table; this replaces the
// sync entry with one that scans the update first (see `update-scan.ts`) and drops a message over
// budget before Yjs decodes it. New code.
import * as decoding from 'lib0/decoding';

import { MAX_UPDATE_BYTES } from './codec';
import { scanUpdate, type UpdateLimits } from './update-scan';

const MESSAGE_SYNC = 0;
const SYNC_STEP_2 = 1;
const SYNC_UPDATE = 2;

type MessageHandler = (encoder: unknown, decoder: decoding.Decoder, ...rest: unknown[]) => void;

/** A y-websocket `WebsocketProvider`, structurally: its per-instance message handler table. */
export interface MessageHandlerProviderLike {
	messageHandlers?: (MessageHandler | undefined)[];
}

export interface ProviderGuardOptions extends UpdateLimits {
	/** Largest update accepted. Default {@link MAX_UPDATE_BYTES}. */
	maxBytes?: number;
	/** Told about each dropped message. */
	onRejected?: (reason: string) => void;
}

/**
 * Make `provider` scan each incoming update before it is decoded, dropping one that is too large or
 * would allocate more than `options` allow. Returns false when the provider has no handler table to
 * guard (it is not a y-websocket provider), in which case nothing changed.
 */
export function guardProviderUpdates(
	provider: MessageHandlerProviderLike,
	options: ProviderGuardOptions = {},
): boolean {
	const handlers = provider.messageHandlers;
	const sync = handlers?.[MESSAGE_SYNC];
	if (!handlers || typeof sync !== 'function') return false;
	const maxBytes = options.maxBytes ?? MAX_UPDATE_BYTES;
	handlers[MESSAGE_SYNC] = (encoder, decoder, ...rest) => {
		const peek = decoding.clone(decoder);
		let reason: string | undefined;
		try {
			const kind = decoding.readVarUint(peek);
			if (kind === SYNC_STEP_2 || kind === SYNC_UPDATE) {
				const update = decoding.readVarUint8Array(peek);
				if (update.byteLength > maxBytes) reason = `update exceeds the ${maxBytes} byte limit`;
				else {
					const scan = scanUpdate(update, options);
					if (!scan.ok) reason = scan.reason;
				}
			}
		} catch {
			reason = 'malformed sync message';
		}
		if (reason !== undefined) {
			options.onRejected?.(reason);
			return;
		}
		sync(encoder, decoder, ...rest);
	};
	return true;
}
