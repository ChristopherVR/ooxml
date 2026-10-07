// The contract every framework binding implements: the same props in, the same events out. The
// element bindings are lifecycle adapters only (create the element, forward props, re-emit events);
// behaviour lives in <teams-app>. Each binding also ships raw hooks over the core client, for apps
// that want to render their own UI (see `store.ts`).
import type { TeamsServerConfig } from 'ooxml-core/teams';
import type { FileOpeners, FileUploader, OpenFileDetail, TeamsApp } from './teams-app';
import type { FileEmbeds } from './content-preview';

export interface TeamsProps {
	/** Shared-state room: 1-100 alphanumeric, `-` or `_`. Default `demo`. */
	workspaceId?: string;
	userName?: string;
	/** Stable id for authorship. Defaults to one remembered in this browser. */
	userId?: string;
	/** Server settings. Omit to use the settings dialog or local mode. */
	config?: TeamsServerConfig | null;
	uploadFile?: FileUploader;
	openers?: FileOpeners;
	embeds?: FileEmbeds;
	onReady?: (detail: { user: { id: string; name: string } }) => void;
	/** Cancel the event (`preventDefault`) to take over opening a file. */
	onOpenFile?: (detail: OpenFileDetail, event: CustomEvent<OpenFileDetail>) => void;
	onConfigChange?: (detail: { config: TeamsServerConfig }) => void;
}

const SCALARS = ['workspaceId', 'userName', 'userId'] as const;

/** Push props onto the element. Only changed values are written, so frameworks can call it freely. */
export function applyTeamsProps(el: TeamsApp, props: TeamsProps, prev: TeamsProps = {}): void {
	for (const key of SCALARS) {
		if (props[key] === prev[key]) continue;
		el[key] = props[key] ?? (key === 'workspaceId' ? 'demo' : '');
	}
	if (props.config !== prev.config && props.config !== undefined) el.config = props.config;
	if (props.uploadFile !== prev.uploadFile) el.uploadFile = props.uploadFile;
	if (props.openers !== prev.openers) el.openers = props.openers ?? {};
	if (props.embeds !== prev.embeds) el.embeds = props.embeds ?? {};
}

/** Subscribe to the element's events; the latest handlers are read through `get`. */
export function listenTeamsEvents(el: TeamsApp, get: () => TeamsProps): () => void {
	const ready = (e: Event): void => get().onReady?.((e as CustomEvent).detail);
	const open = (e: Event): void => {
		const ce = e as CustomEvent<OpenFileDetail>;
		get().onOpenFile?.(ce.detail, ce);
	};
	const config = (e: Event): void => get().onConfigChange?.((e as CustomEvent).detail);
	el.addEventListener('teams-ready', ready);
	el.addEventListener('teams-open-file', open);
	el.addEventListener('teams-config-change', config);
	return () => {
		el.removeEventListener('teams-ready', ready);
		el.removeEventListener('teams-open-file', open);
		el.removeEventListener('teams-config-change', config);
	};
}
