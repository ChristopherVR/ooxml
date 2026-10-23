// The contract every framework binding implements: the same props in, the same events out. The
// element bindings are lifecycle adapters only (create the element, forward props, re-emit events);
// behaviour lives in <teams-app>. Each binding also ships raw hooks over the core client, for apps
// that want to render their own UI (see `store.ts`). The wiring is the shared `ooxml-ui/bindings`
// kit; `bindTeams` adds the host class on top of it.
import type { TeamsServerConfig } from 'ooxml-core/teams';
import {
	bindElement,
	eventArgs,
	listenEvents,
	pickProps,
	syncClassName,
	type BindingContract,
	type PropRule,
} from '../../bindings/index';
import type { FileOpeners, FileUploader, OpenFileDetail, TeamsApp } from './teams-app';
import type { FileEmbeds } from './content-preview';

export interface TeamsProps {
	/** Shared-state room: 1-100 alphanumeric, `-` or `_`. Default `demo`. */
	workspaceId?: string;
	userName?: string;
	/** Stable id for authorship. Defaults to one remembered in this browser. */
	userId?: string;
	/** Display language: a BCP 47 tag such as `de` or `zh-CN`. English, German, Spanish, French and Simplified Chinese ship; empty or unknown means English. */
	locale?: string;
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

type TeamsPropKey =
	| 'workspaceId'
	| 'userName'
	| 'userId'
	| 'locale'
	| 'config'
	| 'uploadFile'
	| 'openers'
	| 'embeds';

/** The props, events and (empty) handle of every Teams binding; the element is the handle. */
export const TEAMS_CONTRACT: BindingContract<
	TeamsApp,
	Pick<TeamsProps, TeamsPropKey>,
	Record<never, never>,
	'teams-ready' | 'teams-open-file' | 'teams-config-change',
	'onReady' | 'onOpenFile' | 'onConfigChange'
> = {
	props: {
		workspaceId: { forward: 'changed', fallback: 'demo' },
		userName: { forward: 'changed', fallback: '' },
		userId: { forward: 'changed', fallback: '' },
		locale: { forward: 'changed', fallback: '' },
		// `null` clears the server settings; an omitted config keeps the dialog's or local mode.
		config: { forward: 'defined' },
		uploadFile: { forward: 'changed' },
		openers: { forward: 'changed', apply: (el, value) => (el.openers = value ?? {}) },
		embeds: { forward: 'changed', apply: (el, value) => (el.embeds = value ?? {}) },
	},
	events: {
		'teams-ready': { option: 'onReady' },
		'teams-open-file': { option: 'onOpenFile', args: eventArgs.detailAndEvent },
		'teams-config-change': { option: 'onConfigChange' },
	},
	handle: {},
};

/** Push props onto the element. Only changed values are written, so frameworks can call it freely. */
export function applyTeamsProps(el: TeamsApp, props: TeamsProps, prev: TeamsProps = {}): void {
	const next = props as Record<string, unknown>;
	const before = prev as Record<string, unknown>;
	for (const [key, rule] of Object.entries(TEAMS_CONTRACT.props) as [
		string,
		PropRule<TeamsApp>,
	][]) {
		const value = next[key];
		if (value === before[key] || (rule.forward === 'defined' && value === undefined)) continue;
		const written = value === undefined && rule.fallback !== undefined ? rule.fallback : value;
		if (rule.apply) rule.apply(el, written);
		else (el as unknown as Record<string, unknown>)[key] = written;
	}
}

/** Subscribe to the element's events; the latest handlers are read through `get`. */
export function listenTeamsEvents(el: TeamsApp, get: () => TeamsProps): () => void {
	const unlisten = listenEvents(el, TEAMS_CONTRACT.events, get);
	return () => unlisten();
}

/** The props of `<teams-app>` plus the class list of its host element. */
export type TeamsElementProps = TeamsProps & { className?: string | undefined };

/**
 * Copies the defined, shared props out of a framework's props/instance object. `undefined` means
 * "not provided"; `config` keeps an explicit `null`. Handlers are not copied: events go through
 * the `get` callback of `bindTeams`.
 */
export function pickTeamsProps(
	source: { [K in TeamsPropKey]?: TeamsProps[K] | undefined } & { className?: string | undefined },
): TeamsElementProps {
	const picked = pickProps<TeamsElementProps>(TEAMS_CONTRACT, source, true);
	if (source.className) picked.className = source.className;
	return picked;
}

export interface TeamsBinding {
	/** Apply props; only values that changed since the last call are written. */
	update(next: TeamsElementProps): void;
	/** Stop listening for events. The element is left to the framework that owns it. */
	destroy(): void;
}

/**
 * Wires `el` to a framework's props: `update` pushes changed props (and the host class), and the
 * element's events call the handlers returned by `get` at event time.
 */
export function bindTeams(el: TeamsApp, get: () => TeamsElementProps): TeamsBinding {
	let className: string | undefined;
	const bound = bindElement(el, TEAMS_CONTRACT, get);
	return {
		update(next) {
			bound.update(next);
			syncClassName(el, className, next.className);
			className = next.className;
		},
		destroy: () => bound.unlisten(),
	};
}
