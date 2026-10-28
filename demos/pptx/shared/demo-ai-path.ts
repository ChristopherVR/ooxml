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
