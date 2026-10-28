/**
 * The chat panel loader a `createAiAssistant` handle (`pptx-svelte-viewer/ai`) carries.
 * `AiDock` calls it the first time the user opens the assistant; without a handle it runs
 * its own `import('./AiChatPanel.svelte')`.
 */
export type AiChatPanelLoader = () => Promise<typeof import('./AiChatPanel.svelte')>;
