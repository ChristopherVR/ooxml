/**
 * useAiConversation: the entire client tool-loop glue between the shared
 * {@link PptxAiChatSession} and `@ai-sdk/vue`'s `useChat`.
 *
 * `useChat` owns the message list, streaming status, and transport request. On
 * each tool call we run the tool against the deck via `session.executeToolCall`
 * and hand the JSON result (or error) back with `addToolOutput`; the session's
 * `sendAutomaticallyWhen` predicate then resubmits so the model can continue.
 *
 * Everything reachable from here (including `@ai-sdk/vue`) sits inside the
 * lazily-loaded AI chat-panel chunk, so importing the viewer never forces the
 * optional SDK peers to be installed.
 *
 * `useChat` is INJECTED (see {@link UseChatFn}), not statically imported: that
 * package is an optional peer, so a consumer who has not installed it gets an
 * empty stub module from their bundler's optional-peer handling. A static
 * `import { useChat } from '@ai-sdk/vue'` asks Rollup to validate that named
 * binding at link time, which fails the CONSUMER's production build outright
 * even though this module is only reached once the AI panel actually opens.
 * `AiConversation.vue` resolves the real `useChat` with a top-level
 * `await import('@ai-sdk/vue')` (Vue's compiler wraps it with
 * `withAsyncContext`, so the reactive setup context survives the await) and
 * passes it in here, so this module never has to import the SDK itself (see
 * issue #143, fixed for `pptx-svelte-viewer`; this is the same defect in
 * `pptx-vue-viewer`).
 */
import type { ChatStatus, UIMessage } from 'ai';
import type {
	PptxAiChatSession,
	PptxAiConfig,
	ProposalView,
	ToolCanvasTarget,
} from 'ooxml-ui/pptx/ai';
import { toolCanvasTarget } from 'ooxml-ui/pptx/ai';
import { computed, ref, watch } from 'vue';
import type { ComputedRef, Ref } from 'vue';

import { extractReadyToolCalls } from './message-parts';

/** The tool call surfaced by `useChat`'s `onToolCall` (narrowed for our tools). */
interface IncomingToolCall {
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

export interface UseAiConversationResult {
	messages: Ref<UIMessage[]>;
	status: Ref<ChatStatus>;
	error: Ref<Error | undefined>;
	isStreaming: ComputedRef<boolean>;
	proposals: Ref<ProposalView[]>;
	/** Loading another presentation dropped the staged proposals (`ProposalStore.discarded`). */
	proposalsDiscarded: Ref<boolean>;
	send: (text: string) => void;
	stop: () => void;
	clearError: () => void;
	applyProposal: (id: string) => void;
	rejectProposal: (id: string) => void;
	acceptAllProposals: () => void;
}

/** What the panel wires in: the deck load, and hooks such as the live on-canvas focus. */
export interface UseAiConversationOptions {
	/** The bridge's deck load: proposals are re-read when it changes, since a load drops them. */
	deckLoad: () => number;
	/**
	 * Called as each tool call starts, with the slide / element(s) it references
	 * (or `null` for deck-wide tools). Lets the panel navigate + highlight the
	 * canvas so the viewer reflects what the assistant is doing in real time.
	 */
	onToolTarget?: (target: ToolCanvasTarget | null) => void;
}

/** `@ai-sdk/vue`'s `useChat`, resolved at runtime by `AiConversation.vue` (see file doc). */
export type UseChatFn = typeof import('@ai-sdk/vue').useChat;

export function useAiConversation(
	useChat: UseChatFn,
	session: PptxAiChatSession,
	config: PptxAiConfig,
	options: UseAiConversationOptions,
): UseAiConversationResult {
	const proposals = ref<ProposalView[]>(session.proposals.list());
	const proposalsDiscarded = ref(session.proposals.discarded);
	const refreshProposals = (): void => {
		proposals.value = session.proposals.list();
		proposalsDiscarded.value = session.proposals.discarded;
	};
	watch(options.deckLoad, refreshProposals);

	// `useChat` is a parameter, resolved once by `AiConversation.vue` when the
	// optional `@ai-sdk/vue` peer loads (see this file's doc comment above), so it
	// stays referentially stable across calls even though the linter can't see
	// that through a parameter.
	// oxlint-disable-next-line react/hooks -- see comment above
	const clientExecutes = session.clientExecutesTools;
	const chat = useChat({
		transport: session.transport,
		// In `model` mode the session's agent runs the tools; running them here as
		// well would stage every proposal twice.
		sendAutomaticallyWhen: clientExecutes ? session.sendAutomaticallyWhen : undefined,
		onError: (err: Error) => config.onError?.(err),
		onToolCall: clientExecutes
			? async ({ toolCall }: { toolCall: unknown }) => {
					const { toolName, toolCallId, input } = toolCall as IncomingToolCall;
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
					refreshProposals();
				}
			: undefined,
	});

	const isStreaming = computed(
		() => chat.status.value === 'submitted' || chat.status.value === 'streaming',
	);

	// Live "AI as a collaborator" focus, driven from the message stream so it
	// works in EVERY connection mode (in `model` mode the agent runs the tool
	// loop and `onToolCall` never fires client-side). Each tool call fires the
	// focus callback exactly once, the moment its input is available.
	const seenToolCalls = new Set<string>();
	watch(
		chat.messages,
		(messages) => {
			// The agent stages proposals while it streams, so nothing else refreshes them.
			refreshProposals();
			for (const call of extractReadyToolCalls(messages)) {
				if (seenToolCalls.has(call.toolCallId)) {
					continue;
				}
				seenToolCalls.add(call.toolCallId);
				options.onToolTarget?.(toolCanvasTarget(call.toolName, call.input));
			}
		},
		{ deep: true },
	);

	const send = (text: string): void => {
		const trimmed = text.trim();
		if (trimmed.length === 0) {
			return;
		}
		void chat.sendMessage({ text: trimmed });
	};

	const applyProposal = (id: string): void => {
		session.proposals.apply(id);
		refreshProposals();
	};
	const rejectProposal = (id: string): void => {
		session.proposals.revert(id);
		refreshProposals();
	};
	const acceptAllProposals = (): void => {
		session.proposals.acceptAll();
		refreshProposals();
	};

	return {
		messages: chat.messages,
		status: chat.status,
		error: chat.error,
		isStreaming,
		proposals,
		proposalsDiscarded,
		send,
		stop: () => void chat.stop(),
		clearError: () => chat.clearError(),
		applyProposal,
		rejectProposal,
		acceptAllProposals,
	};
}
