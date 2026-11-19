/**
 * The transport every binding talks to. It wraps whatever the connection
 * resolved to and adds what must hold for all of them: a fresh deck and focus
 * context on each request, the run limits, and one normalized finish reason.
 *
 * Only the copy of the messages sent to the model carries the context, so the
 * visible transcript stays clean and context never accumulates across turns.
 */

import type { ChatTransport, UIMessageChunk } from 'ai';

import type { PptxAiBridge } from './bridge';
import type { PptxAiContextStrategy, PptxAiUIMessage } from './config';
import { buildDeckContext } from './focused-context';
import type { PptxAiFinishReason, PptxAiRunFinish, RunLimiter } from './run-limits';
import { RunLimitError } from './run-limits';

type SendOptions = Parameters<ChatTransport<PptxAiUIMessage>['sendMessages']>[0];

const FINISH_REASONS: Record<string, PptxAiFinishReason> = {
	stop: 'stop',
	'tool-calls': 'tool-calls',
	length: 'length',
	'content-filter': 'filtered',
	error: 'error',
	interrupted: 'cancelled',
	other: 'unknown',
};

const ENDING_NOTICES: Partial<Record<PptxAiFinishReason, string>> = {
	length: 'The reply was cut off because it reached the length limit.',
	filtered: 'The provider stopped this reply with its content filter.',
	unknown: 'The reply ended unexpectedly, so it may be incomplete.',
};

/** Prepend `context` to the last user message of a copy of `messages`. */
export function injectDeckContext(messages: PptxAiUIMessage[], context: string): PptxAiUIMessage[] {
	const index = messages.map((message) => message.role).lastIndexOf('user');
	if (index < 0) {
		return messages;
	}
	const target = messages[index];
	const next = messages.slice();
	next[index] = {
		...target,
		parts: [
			{
				type: 'text',
				text: `Current deck context (for your reference; do not repeat it back verbatim):\n\n${context}`,
			},
			...target.parts,
		],
	};
	return next;
}

async function withDeckContext(
	messages: PptxAiUIMessage[],
	bridge: PptxAiBridge,
	strategy: PptxAiContextStrategy,
): Promise<PptxAiUIMessage[]> {
	if (strategy === 'none' && !bridge.getFocusedTargets) {
		return messages;
	}
	try {
		const context = await buildDeckContext(bridge, { strategy, focus: {} });
		return context.trim() ? injectDeckContext(messages, context) : messages;
	} catch {
		return messages;
	}
}

export interface SessionTransportOptions {
	bridge: PptxAiBridge;
	contextStrategy: PptxAiContextStrategy;
	limiter: RunLimiter;
	/** The connection runs the tool loop itself, so ending on a tool call means it hit its step limit. */
	agentLoop?: boolean;
	onFinish(finish: PptxAiRunFinish): void;
}

/** A stream that only ends the run with `message`. */
function endingStream(message: string): ReadableStream<UIMessageChunk> {
	return new ReadableStream({
		start(controller) {
			controller.enqueue({ type: 'error', errorText: message });
			controller.close();
		},
	});
}

export function createSessionTransport(
	inner: ChatTransport<PptxAiUIMessage>,
	options: SessionTransportOptions,
): ChatTransport<PptxAiUIMessage> {
	const { bridge, contextStrategy, limiter, onFinish } = options;

	/** Report a tripped limit, returning the notice to show. */
	const limitNotice = (): string | undefined => {
		const limit = limiter.tripped;
		if (!limit) {
			return undefined;
		}
		onFinish({ reason: 'limit', limit });
		return new RunLimitError(limit).message;
	};

	function watch(
		stream: ReadableStream<UIMessageChunk>,
		request: SendOptions,
		timedOut: Promise<void>,
		release: () => void,
	): ReadableStream<UIMessageChunk> {
		const reader = stream.getReader();
		let finished = false;

		/** End the stream, with a notice when the run did not end normally. */
		const close = (
			controller: ReadableStreamDefaultController<UIMessageChunk>,
			ending: PptxAiFinishReason,
		): void => {
			release();
			const notice = limitNotice();
			if (notice) {
				controller.enqueue({ type: 'error', errorText: notice });
			} else {
				onFinish({ reason: ending });
				const text = ENDING_NOTICES[ending];
				if (text) {
					controller.enqueue({ type: 'error', errorText: text });
				}
			}
			controller.close();
		};

		return new ReadableStream<UIMessageChunk>({
			async pull(controller) {
				let result: ReadableStreamReadResult<UIMessageChunk> | 'timeout';
				try {
					result = await Promise.race([reader.read(), timedOut.then(() => 'timeout' as const)]);
				} catch (error) {
					if (limiter.tripped) {
						close(controller, 'error');
					} else if (request.abortSignal?.aborted) {
						release();
						onFinish({ reason: 'cancelled' });
						controller.close();
					} else {
						release();
						onFinish({ reason: 'error' });
						controller.error(error);
					}
					return;
				}
				if (result === 'timeout') {
					void reader.cancel().catch(() => undefined);
					close(controller, 'error');
					return;
				}
				if (result.done) {
					if (finished) {
						release();
						controller.close();
					} else if (request.abortSignal?.aborted) {
						release();
						onFinish({ reason: 'cancelled' });
						controller.close();
					} else {
						close(controller, 'unknown');
					}
					return;
				}
				const chunk = result.value;
				if (chunk.type === 'error' || chunk.type === 'abort') {
					onFinish({ reason: chunk.type === 'error' ? 'error' : 'cancelled' });
					finished = true;
				} else if (chunk.type === 'finish') {
					finished = true;
					const ending = chunk.finishReason
						? (FINISH_REASONS[chunk.finishReason] ?? 'unknown')
						: 'stop';
					if (options.agentLoop && ending === 'tool-calls') {
						limiter.trip('steps');
					}
					if (limiter.tripped || ENDING_NOTICES[ending]) {
						void reader.cancel().catch(() => undefined);
						close(controller, ending);
						return;
					}
					onFinish({ reason: ending });
				}
				controller.enqueue(chunk);
			},
			cancel(reason) {
				release();
				onFinish({ reason: 'cancelled' });
				return reader.cancel(reason);
			},
		});
	}

	return {
		async sendMessages(request) {
			// A request after a tool result continues the run; a restored chat may have none yet.
			if (request.messages.at(-1)?.role === 'user' || !limiter.started) {
				limiter.begin();
			}
			if (!limiter.admitStep()) {
				return endingStream(limitNotice() ?? '');
			}
			const messages = await withDeckContext(request.messages, bridge, contextStrategy);

			const clock = new AbortController();
			let expire: () => void = () => undefined;
			const timedOut = new Promise<void>((resolve) => {
				expire = resolve;
			});
			const timer = setTimeout(() => {
				limiter.trip('time');
				clock.abort();
				expire();
			}, limiter.remainingMs());
			const release = () => clearTimeout(timer);
			const abortSignal = request.abortSignal
				? AbortSignal.any([request.abortSignal, clock.signal])
				: clock.signal;
			try {
				const stream = await inner.sendMessages({ ...request, messages, abortSignal });
				return watch(stream, request, timedOut, release);
			} catch (error) {
				release();
				const notice = limitNotice();
				if (notice) {
					return endingStream(notice);
				}
				onFinish({ reason: request.abortSignal?.aborted ? 'cancelled' : 'error' });
				throw error;
			}
		},
		reconnectToStream: (request) => inner.reconnectToStream(request),
	};
}
