/**
 * Which value every demo hands its viewer's `ai` prop, input or option, resolved from the
 * page URL so all five demos agree:
 *
 *  - By default, the handle `createAiAssistant(config)` returns from the binding's `/ai`
 *    entry, the path the AI assistant guide recommends.
 *  - With `?ai=config`, the plain `PptxAiConfig`, which stays supported. The browser suite
 *    (`e2e/pptx/ai-panel-parity.spec.ts`) drives both paths in every binding.
 *
 * Each demo asks only when it has a config to pass, and the answer is recorded on
 * `document.documentElement.dataset.demoAi` (`'config'` or `'handle'`), so the browser
 * suite can check which value the viewer got.
 */
export function demoAiUsesPlainConfig(search: string = globalThis.location?.search ?? ''): boolean {
	const plain = new URLSearchParams(search).get('ai') === 'config';
	if (typeof document !== 'undefined') {
		document.documentElement.dataset.demoAi = plain ? 'config' : 'handle';
	}
	return plain;
}

/** The assistant write policies a demo can run under. */
export type DemoAiWritePolicy = 'auto' | 'stage' | 'approve';

/**
 * The `writePolicy` every demo passes, from `?aiPolicy=auto` or `?aiPolicy=approve`. Anything else
 * is `'stage'`, the library's default, so the browser suite can drive each policy through the
 * real panel (`e2e/pptx/ai-write-policy.spec.ts`).
 */
export function demoAiWritePolicy(
	search: string = globalThis.location?.search ?? '',
): DemoAiWritePolicy {
	const policy = new URLSearchParams(search).get('aiPolicy');
	return policy === 'auto' || policy === 'approve' ? policy : 'stage';
}
