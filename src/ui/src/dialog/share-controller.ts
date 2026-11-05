import type { CollaborationOptions, CollaborationState, ShareFields } from 'ooxml-core/collab';
import { type OfficeShareDialogLabels, shareDialogLabels } from './share-dialog-locales';
import { defineShareDialog, type OfficeUiShareDialog } from './share-dialog';

/** What the Share dialog needs from a product editor: its session, as `CollaborativeEditor` has it. */
export interface ShareDialogHost {
	state(): CollaborationState;
	/** Start or join; a rejection (or throw) is shown in the dialog. */
	start(options: CollaborationOptions): Promise<void> | void;
	stop(): void;
}

export interface ShareDialogSetup {
	/** Where the dialog element goes (the editor's shadow root or frame). */
	container: ParentNode;
	host: ShareDialogHost;
	/** A suggested session name starts with this (`doc`, `book`, `drawing`). */
	roomPrefix: string;
	/** The form's defaults each time the dialog opens: the user's name, the host's server. */
	defaults?: () => Partial<ShareFields>;
	/** The product's own line on what is shared, already translated. */
	note?: (locale: string | undefined) => string;
	locale?: () => string | undefined;
}

/** One product's Share dialog: open it, keep it in step with the session, translate it. */
export interface ShareDialogController {
	readonly element: OfficeUiShareDialog;
	/** Open it, on the join tab when `invitation` is a link or a session name. */
	open(invitation?: string): void;
	/** Re-read the session (call on `collaboration-change`). */
	refresh(): void;
	relocalize(): void;
	destroy(): void;
}

const message = (cause: unknown): string =>
	cause instanceof Error ? cause.message : String(cause);

/**
 * Mount the shared `office-ui-share-dialog` for a product editor. Start and Join hand the dialog's
 * `CollaborationOptions` to `host.start` (the editor's `collaboration`), Stop calls `host.stop`,
 * and the dialog follows `host.state()` on every `refresh`. Word, Excel and Visio all use it, so
 * their Share buttons share, join and invite the same way PowerPoint's does.
 */
export function attachShareDialog(setup: ShareDialogSetup): ShareDialogController {
	const doc = (setup.container as Node).ownerDocument ?? (setup.container as Document);
	defineShareDialog(doc.defaultView?.customElements);
	const element = doc.createElement('office-ui-share-dialog') as OfficeUiShareDialog;
	element.roomPrefix = setup.roomPrefix;
	const relocalize = (): void => {
		const locale = setup.locale?.();
		const labels: OfficeShareDialogLabels = shareDialogLabels(locale);
		element.labels = labels;
		element.note = setup.note?.(locale) ?? '';
	};
	const refresh = (): void => {
		element.state = setup.host.state();
	};
	element.addEventListener('office-share-start', (event) => {
		const { options } = (event as CustomEvent<{ options: CollaborationOptions }>).detail;
		element.error = '';
		try {
			Promise.resolve(setup.host.start(options))
				.then(refresh)
				.catch((cause: unknown) => {
					element.error = message(cause);
					refresh();
				});
		} catch (cause) {
			element.error = message(cause);
		}
		refresh();
	});
	element.addEventListener('office-share-stop', () => {
		element.error = '';
		try {
			setup.host.stop();
		} catch (cause) {
			element.error = message(cause);
		}
		refresh();
	});
	setup.container.append(element);
	relocalize();
	refresh();
	return {
		element,
		open(invitation = '') {
			relocalize();
			refresh();
			element.defaults = setup.defaults?.() ?? {};
			element.invitation = invitation;
			element.error = '';
			element.open = true;
		},
		refresh,
		relocalize,
		destroy() {
			element.remove();
		},
	};
}
