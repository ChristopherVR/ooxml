/**
 * `SvelteAiChat` - the runes-based chat controller behind the AI panel.
 *
 * It guards on {@link isAiAvailable}, builds the framework-agnostic
 * {@link PptxAiChatSession} for the given bridge + config, and drives
 * `@ai-sdk/svelte`'s runes-based {@link Chat}. Message/status/error come
 * straight off `Chat`'s reactive getters; staged proposals live in the
 * session's `ProposalStore` (outside the chat stream), so they are mirrored
 * into a `$state` array and refreshed after each tool call / accept / reject.
 *
 * This module (and, through it, `@ai-sdk/svelte` + the optional `ai` SDK) is
 * only reached through the lazily-imported `AiChatPanel.svelte`, so the base
 * viewer bundle never statically pulls the SDK.
 *
 * `Chat` itself is loaded with a runtime `import()`, not a static named
 * import: `@ai-sdk/svelte` is an optional peer, and a consumer who has not
 * installed it gets an empty stub module from their bundler's optional-peer
 * handling. A static `import { Chat } from '@ai-sdk/svelte'` asks Rollup to
 * validate that named binding at link time, which fails the CONSUMER's
 * production build outright (`"Chat" is not exported`) even though this
 * module is never reached unless the AI panel actually opens. A dynamic
 * `import()` defers that lookup to when `Chat` is actually read off the
 * resolved module, at runtime, inside `init()`, so a consumer who never
 * installs the SDK can still build (see issue #143).
 */

import type { Chat as ChatType } from '@ai-sdk/svelte';
import type { ChatStatus } from 'ai';
import {
	createAiChatSession,
	extractReadyToolCalls,
	isAiAvailable,
	toolCanvasTarget,
} from 'ooxml-ui/pptx/ai';
import type {
	PptxAiBridge,
	PptxAiChatSession,
	PptxAiConfig,
	PptxAiUIMessage,
	ProposalView,
	ToolCanvasTarget,
} from 'ooxml-ui/pptx/ai';

/** Lifecycle of the AI session bootstrap. */
export type AiChatInitState = 'checking' | 'unavailable' | 'ready' | 'error';

/** The tool call surfaced by `Chat`'s `onToolCall` (narrowed for our tools). */
interface ToolCallInfo {
	toolName: string;
	toolCallId: string;
	input: unknown;
}

/** Loosely-typed `addToolOutput` (the SDK types `tool` as a tool-name union). */
type AddToolOutput = (arg: {
	tool: string;
	toolCallId: string;
	output?: unknown;
	state?: 'output-error';
	errorText?: string;
}) => void;

/** Inputs + injectable seams (tests replace the async builders). */
export interface SvelteAiChatDeps {
	bridge: PptxAiBridge;
	config: PptxAiConfig;
	/** Availability probe; defaults to {@link isAiAvailable}. */
	checkAvailable?: () => Promise<boolean>;
	/** Session builder; defaults to {@link createAiChatSession}. */
	createSession?: (bridge: PptxAiBridge, config: PptxAiConfig) => Promise<PptxAiChatSession>;
	/**
	 * Called as each tool call starts, with the slide / element(s) it references
	 * (or `null` for deck-wide tools). Lets the panel navigate + highlight the
	 * canvas so the viewer mirrors what the assistant is doing in real time.
	 */
	onToolTarget?: (target: ToolCanvasTarget | null) => void;
}

export class SvelteAiChat {
	initState = $state<AiChatInitState>('checking');
	initError = $state<Error | undefined>(undefined);
	/** Mirror of the session's staged proposals (refreshed imperatively). */
	proposals = $state<ProposalView[]>([]);
	/** Loading another presentation dropped the staged proposals (see `ProposalStore.discarded`). */
	proposalsDiscarded = $state(false);

	#chat = $state.raw<ChatType<PptxAiUIMessage> | undefined>(undefined);
	// Reactive so a consumer `$effect` (the panel's change-animator subscription)
	// re-runs once the session is bootstrapped by `init()`.
	#session = $state.raw<PptxAiChatSession | undefined>(undefined);
	readonly #deps: SvelteAiChatDeps;
	#seenToolCalls = new Set<string>();

	constructor(deps: SvelteAiChatDeps) {
		this.#deps = deps;
	}

	get messages(): PptxAiUIMessage[] {
		return (this.#chat?.messages ?? []) as PptxAiUIMessage[];
	}

	get status(): ChatStatus {
		return this.#chat?.status ?? 'ready';
	}

	get error(): Error | undefined {
		return this.#chat?.error;
	}

	get isStreaming(): boolean {
		return this.status === 'submitted' || this.status === 'streaming';
	}

	/** The live session, exposed for tests that stage a proposal directly. */
	get session(): PptxAiChatSession | undefined {
		return this.#session;
	}

	/** Bootstrap the session + chat. Safe to call once (from the panel `$effect`). */
	async init(): Promise<void> {
		this.initState = 'checking';
		this.initError = undefined;
		try {
			const available = await (this.#deps.checkAvailable ?? isAiAvailable)();
			if (!available) {
				this.initState = 'unavailable';
				return;
			}
			const session = await (this.#deps.createSession ?? createAiChatSession)(
				this.#deps.bridge,
				this.#deps.config,
			);
			this.#session = session;
			const { Chat } = await import('@ai-sdk/svelte');
			this.#chat = new Chat<PptxAiUIMessage>({
				transport: session.transport,
				onError: (err: Error) => this.#deps.config.onError?.(err),
				// In `model` mode the session's agent runs the tools; running them here
				// as well would stage every proposal twice.
				...(session.clientExecutesTools
					? {
							sendAutomaticallyWhen: session.sendAutomaticallyWhen,
							onToolCall: async ({ toolCall }: { toolCall: unknown }) => {
								const { toolName, toolCallId, input } = toolCall as ToolCallInfo;
								const chat = this.#chat;
								if (!chat) {
									return;
								}
								const addToolOutput = chat.addToolOutput as unknown as AddToolOutput;
								try {
									const output = await session.executeToolCall(toolName, input);
									addToolOutput({ tool: toolName, toolCallId, output });
								} catch (err) {
									addToolOutput({
										tool: toolName,
										toolCallId,
										state: 'output-error',
										errorText: err instanceof Error ? err.message : String(err),
									});
								}
								// A staged proposal was likely just registered; surface it.
								this.refreshProposals();
							},
						}
					: {}),
			});
			this.initState = 'ready';
		} catch (err) {
			this.initError = err instanceof Error ? err : new Error(String(err));
			this.initState = 'error';
		}
	}

	send(text: string): void {
		const trimmed = text.trim();
		if (trimmed.length === 0 || !this.#chat) {
			return;
		}
		void this.#chat.sendMessage({ text: trimmed });
	}

	stop(): void {
		void this.#chat?.stop();
	}

	/**
	 * Follow the message stream, which works in every connection mode: report each
	 * new tool call to `onToolTarget` once, as soon as its input is known, and show
	 * the proposals the agent has staged. The panel runs this from an effect, which
	 * also re-runs on every presentation load: listing reads the bridge's deck load,
	 * so proposals a load dropped leave the panel at once.
	 */
	followToolCalls(): void {
		this.refreshProposals();
		for (const call of extractReadyToolCalls(this.messages)) {
			if (!this.#seenToolCalls.has(call.toolCallId)) {
				this.#seenToolCalls.add(call.toolCallId);
				this.#deps.onToolTarget?.(toolCanvasTarget(call.toolName, call.input));
			}
		}
	}

	/** Replace the whole transcript (chat-history resume / new chat). */
	setMessages(messages: PptxAiUIMessage[]): void {
		// A resumed chat's finished tool calls must not move the canvas.
		for (const call of extractReadyToolCalls(messages)) {
			this.#seenToolCalls.add(call.toolCallId);
		}
		if (this.#chat) {
			this.#chat.messages = messages;
		}
	}

	clearError(): void {
		this.#chat?.clearError();
	}

	refreshProposals(): void {
		this.proposals = this.#session?.proposals.list() ?? [];
		this.proposalsDiscarded = this.#session?.proposals.discarded ?? false;
	}

	applyProposal(id: string): void {
		this.#session?.proposals.apply(id);
		this.refreshProposals();
	}

	rejectProposal(id: string): void {
		this.#session?.proposals.revert(id);
		this.refreshProposals();
	}

	acceptAllProposals(): void {
		this.#session?.proposals.acceptAll();
		this.refreshProposals();
	}
}
