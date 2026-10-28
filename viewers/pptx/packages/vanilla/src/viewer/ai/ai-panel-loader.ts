/**
 * The panel loader a `createAiAssistant` handle (`pptx-vanilla-viewer/ai`) carries.
 * `mountAiChat` calls it the first time the user opens the assistant; without a handle it
 * runs its own `import('./ai-panel')`.
 */
export type AiPanelLoader = () => Promise<typeof import('./ai-panel')>;
