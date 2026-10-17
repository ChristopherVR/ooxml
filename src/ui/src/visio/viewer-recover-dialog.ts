import type { RecoveryOffer } from './viewer-autorecover';

type Dialog = HTMLElement & { open: boolean; show(): void; close(): void };

/**
 * "Recover Drawing": shown when a file is opened that has a kept AutoRecover copy. The file name
 * and time are user data and go in as text only. Restore opens the kept copy; Discard deletes it;
 * closing the dialog keeps it for next time.
 */
export class RecoverDialog {
	readonly dialog: Dialog;
	readonly #message: HTMLElement;
	#offer: RecoveryOffer | undefined;
	constructor(
		root: ShadowRoot,
		/** Told what was done, for the status line. */
		private readonly done: (message: string) => void,
		private readonly locale: () => string | undefined = () => undefined,
	) {
		const doc = root.ownerDocument;
		this.dialog = doc.createElement('office-ui-dialog') as Dialog;
		this.dialog.className = 'recover-dialog';
		this.dialog.setAttribute('heading', 'Recover Drawing');
		this.#message = doc.createElement('p');
		this.dialog.append(this.#message);
		for (const [name, command] of [
			['Keep for later', 'later'],
			['Discard', 'discard'],
			['Restore', 'restore'],
		] as const) {
			const button = doc.createElement('office-ui-button');
			button.slot = 'footer';
			button.setAttribute('label', name);
			button.setAttribute('command', `recover-${command}`);
			button.addEventListener('office-command', () => void this.#choose(command));
			this.dialog.append(button);
		}
		root.append(this.dialog);
	}
	show(offer: RecoveryOffer): void {
		this.#offer = offer;
		const when = new Date(offer.savedAt).toLocaleString(this.locale());
		this.#message.textContent = `A copy of ${offer.name} with unsaved changes was kept in this browser on ${when}. Restore it?`;
		this.dialog.show();
	}
	async #choose(choice: 'later' | 'discard' | 'restore'): Promise<void> {
		const offer = this.#offer;
		this.#offer = undefined;
		this.dialog.close();
		if (!offer || choice === 'later') return;
		try {
			if (choice === 'restore') {
				await offer.restore();
				this.done(`Restored ${offer.name}. Save a copy to keep it.`);
			} else {
				await offer.discard();
				this.done(`Discarded the kept copy of ${offer.name}.`);
			}
		} catch (error) {
			this.done(error instanceof Error ? error.message : 'The kept copy could not be restored.');
		}
	}
}
