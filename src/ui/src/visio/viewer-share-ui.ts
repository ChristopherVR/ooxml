import { attachShareDialog, type ShareDialogController } from '../dialog/share-controller';
import type { OfficeProfile } from '../controls';
import { defineRibbonActions, type OfficeUiRibbonActions } from '../ribbon/ribbon-actions';
import { normalizeVisioLocale, translateRibbonText } from './ribbon-locale';
import { announceFileCommand } from './viewer-file-commands';
import type { ViewerShare } from './viewer-share';

export interface VisioShareUiSetup {
	element: HTMLElement;
	root: ShadowRoot;
	/** The `office-ui-ribbon`; Share goes in its `actions` slot, as in Word and Excel. */
	ribbon: HTMLElement;
	share: ViewerShare;
	profile(): OfficeProfile;
	locale(): string;
}

export interface VisioShareUi {
	readonly dialog: ShareDialogController;
	readonly actions: OfficeUiRibbonActions;
	/** The Share command: announced as `file-command`, then the Share dialog unless a host took it. */
	share(): void;
	refresh(): void;
	relocalize(): void;
	destroy(): void;
}

/**
 * Visio's Share button (the tab row's `office-ui-ribbon-actions`, Share only) and the shared Share
 * dialog, wired to the viewer's session. File > Share > Share with people runs the same command.
 */
export function attachVisioShare(setup: VisioShareUiSetup): VisioShareUi {
	const { element, share } = setup;
	const doc = element.ownerDocument;
	defineRibbonActions(doc.defaultView?.customElements);
	const actions = doc.createElement('office-ui-ribbon-actions') as OfficeUiRibbonActions;
	actions.className = 'ribbon-share';
	actions.slot = 'actions';
	actions.noComments = true;
	setup.ribbon.append(actions);
	const dialog = attachShareDialog({
		container: setup.root,
		host: {
			state: () => share.state,
			start: (options) => share.start(options),
			stop: () => share.stop(),
		},
		roomPrefix: 'drawing',
		defaults: () => ({ userName: setup.profile().displayName }),
		locale: () => setup.locale(),
	});
	const run = (): void => {
		if (announceFileCommand(element, 'share')) dialog.open();
	};
	actions.addEventListener('office-ribbon-share', run);
	share.onOpen = run;
	const relocalize = (): void => {
		const label = translateRibbonText(normalizeVisioLocale(setup.locale()), 'Share');
		actions.shareLabel = label;
		actions.shareTitle = label;
		dialog.relocalize();
	};
	const refresh = (): void => {
		actions.sharePressed = share.state.active;
		dialog.refresh();
	};
	relocalize();
	refresh();
	return {
		dialog,
		actions,
		share: run,
		refresh,
		relocalize,
		destroy() {
			actions.remove();
			dialog.destroy();
		},
	};
}
