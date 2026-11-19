// The format-neutral assistant area. Today it holds the policy a deployment sets (`policy.ts`);
// providers, context packing and the agent runner join it later. DOM-free, and it makes no
// request of its own.
export {
	AI_FEATURES,
	type AiConnection,
	type AiFeature,
	aiFeatureEnabled,
	type AiPolicy,
	type AiProviderSpec,
	checkAiEndpoint,
	normalizeAiPolicy,
	userConnectionRefusal,
} from './policy';
