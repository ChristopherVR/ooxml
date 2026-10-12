import { attachControlStyles } from './control-styles';
export interface PptxUiSlideClipboardElement extends HTMLElement {
	copyAction: (() => Promise<void>) | undefined;
	pasteAction: (() => Promise<void>) | undefined;
	labels: { copy: string; paste: string; copied: string; pasted: string };
}
export function slideClipboardLabels(
	t: (key: string) => string,
): PptxUiSlideClipboardElement['labels'] {
	return {
		copy: t('pptx.slideSorter.systemCopy'),
		paste: t('pptx.slideSorter.systemPaste'),
		copied: t('pptx.slideSorter.systemCopied'),
		pasted: t('pptx.slideSorter.systemPasted'),
	};
}
/** Shared async system-clipboard actions, with visible permission and import failures. */
export function definePptxSlideClipboard(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-slide-clipboard')) return;
	class SlideClipboard extends HTMLElement implements PptxUiSlideClipboardElement {
		private copy: (() => Promise<void>) | undefined;
		private paste: (() => Promise<void>) | undefined;
		private text = {
			copy: 'Copy slides to system clipboard',
			paste: 'Paste slides from system clipboard',
			copied: 'Slides copied',
			pasted: 'Slides pasted',
		};
		private busy = false;
		private readonly copyButton: HTMLButtonElement;
		private readonly pasteButton: HTMLButtonElement;
		private readonly status: HTMLElement;
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachControlStyles(
				root,
				':host{display:inline-flex;align-items:center;flex-wrap:wrap;gap:6px;color:var(--pptx-foreground,var(--foreground,#222))}button{padding:5px 8px;border:1px solid var(--pptx-border,var(--border,#ccc));border-radius:4px;background:var(--pptx-muted,var(--muted,#eee));color:inherit;font:inherit;font-size:12px;cursor:pointer}button:disabled{opacity:.5;cursor:default}[role=status]{max-width:400px;font-size:12px}',
			);
			this.copyButton = document.createElement('button');
			this.copyButton.type = 'button';
			this.pasteButton = document.createElement('button');
			this.pasteButton.type = 'button';
			this.status = document.createElement('span');
			this.status.setAttribute('role', 'status');
			this.status.setAttribute('aria-live', 'polite');
			this.copyButton.addEventListener('click', () => void this.run(this.copy, this.text.copied));
			this.pasteButton.addEventListener('click', () => void this.run(this.paste, this.text.pasted));
			root.append(this.copyButton, this.pasteButton, this.status);
			this.sync();
		}
		get copyAction() {
			return this.copy;
		}
		set copyAction(value: (() => Promise<void>) | undefined) {
			this.copy = value;
			this.sync();
		}
		get pasteAction() {
			return this.paste;
		}
		set pasteAction(value: (() => Promise<void>) | undefined) {
			this.paste = value;
			this.sync();
		}
		get labels() {
			return this.text;
		}
		set labels(value: PptxUiSlideClipboardElement['labels']) {
			this.text = value;
			this.sync();
		}
		private sync(): void {
			this.copyButton.textContent = this.text.copy;
			this.pasteButton.textContent = this.text.paste;
			this.copyButton.disabled = this.busy || !this.copy;
			this.pasteButton.disabled = this.busy || !this.paste;
		}
		private async run(action: (() => Promise<void>) | undefined, success: string): Promise<void> {
			if (!action || this.busy) return;
			this.busy = true;
			this.status.textContent = '';
			this.sync();
			try {
				await action();
				this.status.textContent = success;
			} catch (error) {
				this.status.textContent = error instanceof Error ? error.message : String(error);
			} finally {
				this.busy = false;
				this.sync();
			}
		}
	}
	registry.define('pptx-ui-slide-clipboard', SlideClipboard);
}
