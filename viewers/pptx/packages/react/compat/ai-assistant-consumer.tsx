/**
 * Declaration guard for the `pptx-react-viewer/ai` handle, checked against the BUILT
 * `dist/` declarations by `bun run typecheck:react18` (see `tsconfig.react18.json`).
 *
 * Every entry bundles its own copy of the declarations, so this proves that a handle typed
 * by `/ai` is accepted by the `ai` prop of both `pptx-react-viewer` and
 * `pptx-react-viewer/viewer`, and that the handle stays opaque: it has no members to read,
 * and an object built by hand is not a handle.
 */
import type { PowerPointViewerProps, PptxAiAssistant, PptxAiConfig } from 'pptx-react-viewer';
import { PowerPointViewer } from 'pptx-react-viewer';
import { createAiAssistant } from 'pptx-react-viewer/ai';
import type { PptxAiAssistant as AiEntryAssistant } from 'pptx-react-viewer/ai';
import type { PowerPointViewerProps as ViewerEntryProps } from 'pptx-react-viewer/viewer';
import type { ReactElement } from 'react';

const config: PptxAiConfig = { connection: { kind: 'endpoint', api: '/api/ai/chat' } };
const assistant = createAiAssistant(config);

export const mainEntryAi: PowerPointViewerProps['ai'] = assistant;
export const viewerEntryAi: ViewerEntryProps['ai'] = assistant;
export const plainConfigAi: PowerPointViewerProps['ai'] = config;
export const sameHandleType: PptxAiAssistant = assistant satisfies AiEntryAssistant;

// @ts-expect-error The handle is opaque: it has no `config` to read.
export const unreadable: unknown = assistant.config;

export const handBuilt: PowerPointViewerProps['ai'] = {
	// @ts-expect-error A hand-built object is neither a config nor a handle.
	binding: 'react',
	config,
	panel: () => null,
};

export function AiAssistantConsumer({ content }: { content: Uint8Array }): ReactElement {
	return <PowerPointViewer content={content} ai={assistant} />;
}
