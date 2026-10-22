// The host's AI chat route. The model runs here with the host's key; the deck tools are
// schema-only, so every tool call goes back to the reviewer's browser, runs against the live deck
// and (with the default `stage` policy) waits for the reviewer to accept it. Accepted changes are
// ordinary edits of that reviewer, so they reach the room like typing does.

export interface AiSettings {
	baseURL: string;
	apiKey: string;
	model: string;
}

const LOOPBACK = /^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(:\d+)?\//u;

/** The AI settings from the environment, or null when the host has not configured a model. */
export function aiSettings(
	env: Record<string, string | undefined> = process.env,
): AiSettings | null {
	const model = env.REVIEW_AI_MODEL?.trim();
	const baseURL = (env.REVIEW_AI_BASE_URL?.trim() || 'https://api.openai.com/v1').replace(
		/\/?$/u,
		'/',
	);
	const apiKey = env.REVIEW_AI_API_KEY?.trim() ?? '';
	// A local model (Ollama, LM Studio) needs no key; a remote one does.
	if (!model || (!apiKey && !LOOPBACK.test(baseURL))) return null;
	return { baseURL: baseURL.replace(/\/$/u, ''), apiKey, model };
}

export async function streamReviewChat(
	settings: AiSettings,
	body: unknown,
	context: { reviewTitle: string; reviewer: string },
): Promise<Response> {
	const messages = (body as { messages?: unknown } | null)?.messages;
	if (!Array.isArray(messages)) return new Response('Expected { messages }', { status: 400 });
	const [{ convertToModelMessages, streamText }, { createOpenAICompatible }, deckAi] =
		await Promise.all([
			import('ai'),
			import('@ai-sdk/openai-compatible'),
			import('ooxml-ui/pptx/ai'),
		]);
	const provider = createOpenAICompatible({
		name: 'review-hub',
		baseURL: settings.baseURL,
		apiKey: settings.apiKey,
	});
	const tools = await deckAi.buildPptxAiTools();
	const result = streamText({
		model: provider(settings.model),
		system: deckAi.buildPptxAiSystemPrompt({
			writePolicy: 'stage',
			extras: `You are helping ${context.reviewer} review "${context.reviewTitle}" with colleagues who edit the same deck live. Make small, targeted changes; never restructure or delete slides unless asked.`,
		}),
		tools,
		messages: await convertToModelMessages(messages, { tools }),
	});
	return result.toUIMessageStreamResponse();
}
