/**
 * Demo-only AI wiring: this is the "host supplies the provider" example. The
 * demo reads an OpenAI-compatible base URL, API key, and model id from
 * sessionStorage, then constructs an in-browser language model via
 * `@ai-sdk/openai-compatible` and hands it to the viewer as
 * `ai: { connection: { kind: 'model', model } }`.
 *
 * There is deliberately NO configuration UI on the landing screen: the demo must
 * work, and be shippable, without anyone supplying a key. To try the assistant
 * locally, set the keys by hand (they last until the tab closes):
 *
 *   sessionStorage.setItem('demo.ai.apiKey', 'sk-...')
 *   sessionStorage.setItem('demo.ai.model', 'gpt-4o-mini')
 *
 * Real apps would normally keep the key server-side and use a `kind: 'endpoint'`
 * connection instead; running the model in the browser here keeps the demo
 * self-contained. When the API key or model id is blank the viewer opens with no
 * AI assistant at all.
 */

import { createOpenAICompatible } from '@ai-sdk/openai-compatible';
import type { PptxAiConfig } from 'pptx-vanilla-viewer';
import { createAiAssistant } from 'pptx-vanilla-viewer/ai';
import type { PptxAiAssistant } from 'pptx-vanilla-viewer/ai';

import { demoAiUsesPlainConfig, demoAiWritePolicy } from '../../shared/demo-ai-path';
import { readDemoAiSetting } from '../../shared/demo-ai-storage';

const KEY_BASE_URL = 'demo.ai.baseURL';
const KEY_API_KEY = 'demo.ai.apiKey';
const KEY_MODEL = 'demo.ai.model';

interface AiSettings {
	baseURL: string;
	apiKey: string;
	model: string;
}

function read(): AiSettings {
	return {
		baseURL: readDemoAiSetting(KEY_BASE_URL) ?? 'https://api.openai.com/v1',
		apiKey: readDemoAiSetting(KEY_API_KEY) ?? '',
		model: readDemoAiSetting(KEY_MODEL) ?? 'gpt-4o-mini',
	};
}

/**
 * Build the viewer `ai` option from the tab's demo settings, or `undefined`
 * when the key/model are blank (viewer opens with no assistant).
 */
export function buildViewerAiConfig(): PptxAiConfig | undefined {
	const settings = read();
	if (!settings.apiKey.trim() || !settings.model.trim()) {
		return undefined;
	}
	const provider = createOpenAICompatible({
		name: 'demo-openai-compatible',
		baseURL: settings.baseURL.trim(),
		apiKey: settings.apiKey.trim(),
	});
	return {
		connection: { kind: 'model', model: provider(settings.model.trim()) },
		writePolicy: demoAiWritePolicy(),
	};
}

/**
 * The viewer `ai` value: the `createAiAssistant` handle by default, the plain config with
 * `?ai=config`, or `undefined` when no assistant is configured.
 */
export function buildViewerAi(): PptxAiAssistant | PptxAiConfig | undefined {
	const config = buildViewerAiConfig();
	return config && !demoAiUsesPlainConfig() ? createAiAssistant(config) : config;
}
