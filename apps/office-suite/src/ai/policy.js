/**
 * The assistant's policy: what a deployment allows, decided outside the UI. An organisation puts
 * an `ai-config.json` beside the app (or a host such as the desktop shell sets
 * `globalThis.ooxmlAiConfig` before the app loads) to choose the providers and models, hold the
 * keys behind its own gateway, turn features off, cap context and set the instructions every
 * request carries. The UI only offers what the policy leaves to the user. Without a configuration
 * the assistant is off: no chat, no Ask AI, no assistant pane. The suite demo ships a configuration
 * that turns it on. Normalising and authorising live in `ooxml-core/ai`; this module only loads
 * the configuration. See docs/suite-ai-config.md.
 */
import { aiFeatureEnabled, normalizeAiPolicy } from 'ooxml-core/ai';
import { providers } from './provider-catalog.js';

let policy = normalizeAiPolicy(undefined, providers);
let source = 'default';
let loading;

/** Loads the deployment's policy once: a host-provided object, else `ai-config.json`. */
export function loadPolicy() {
	loading ??= (async () => {
		let raw = globalThis.ooxmlAiConfig;
		if (raw) source = 'host';
		else
			try {
				const root = document.querySelector('meta[name=ooxml-root]')?.content ?? './';
				const response = await fetch(new URL(`${root}ai-config.json`, location.href), {
					cache: 'no-store',
					credentials: 'same-origin',
				});
				if (response.ok && response.headers.get('content-type')?.includes('json')) {
					raw = await response.json();
					// The shipped file says "managed": false; an organisation's own file manages the assistant.
					source = raw?.managed === false ? 'default' : 'file';
				}
			} catch {}
		policy = normalizeAiPolicy(raw, providers);
		return policy;
	})();
	return loading;
}

export const currentPolicy = () => policy;
/** Whether an organisation configured the assistant (so the UI can say who manages it). */
export const isManaged = () => source !== 'default';
export const feature = (name) => aiFeatureEnabled(policy, name);
/** Whether the assistant is on in this deployment. */
export const assistantEnabled = () => policy.enabled;
