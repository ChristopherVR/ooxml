// @vitest-environment jsdom
import { DEFAULT_VIEWER_OPTIONS } from './options/viewer-options';
import {
	alertNoticeDialog,
	confirmActionDialog,
	confirmDeleteDialog,
	confirmExternalHyperlinkDialog,
	openExternalHyperlinkConfirmed,
	promptNameDialog,
} from './viewer-dialogs';

const t = (key: string) => `[${key}]`;
const context = { t };
const dialog = () => document.querySelector('office-ui-message-dialog') as HTMLElement;
const press = (cls: string) =>
	(dialog().shadowRoot!.querySelector(`button.${cls}`) as HTMLButtonElement).click();

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

describe('viewer dialogs', () => {
	it('asks about an external hyperlink with translated text and the address', async () => {
		const answer = confirmExternalHyperlinkDialog('https://example.com/a', context);
		const root = dialog().shadowRoot!;
		expect(root.querySelector('h2')?.textContent).toBe('[pptx.dialog.hyperlinkTitle]');
		expect(root.querySelector('.message')?.textContent).toBe(
			'[pptx.dialog.hyperlinkMessage]\n\nhttps://example.com/a',
		);
		expect(root.querySelector('button.confirm')?.textContent).toBe('[pptx.dialog.hyperlinkOpen]');
		press('confirm');
		expect(await answer).toBe(true);
	});

	it('opens a confirmed link, skips the question when the gate is off, and honours Cancel', async () => {
		const open = vi.spyOn(window, 'open').mockReturnValue(null);
		const gated = structuredClone(DEFAULT_VIEWER_OPTIONS);
		gated.trust.confirmExternalHyperlinks = true;
		const ungated = structuredClone(DEFAULT_VIEWER_OPTIONS);
		ungated.trust.confirmExternalHyperlinks = false;

		const cancelled = openExternalHyperlinkConfirmed(gated, 'https://example.com', context);
		press('cancel');
		expect(await cancelled).toBe(false);
		expect(open).not.toHaveBeenCalled();

		const confirmed = openExternalHyperlinkConfirmed(gated, 'https://example.com', context);
		press('confirm');
		expect(await confirmed).toBe(true);
		expect(open).toHaveBeenCalledOnce();

		expect(await openExternalHyperlinkConfirmed(ungated, 'https://example.org', context)).toBe(
			true,
		);
		expect(dialog()).toBeNull();
		expect(open).toHaveBeenCalledTimes(2);
	});

	it('uses a danger Delete button for destructive confirmations', async () => {
		const answer = confirmDeleteDialog('Delete it?', context);
		expect(dialog().shadowRoot?.querySelector('button.confirm')?.className).toContain('danger');
		expect(dialog().shadowRoot?.querySelector('button.confirm')?.textContent).toBe(
			'[pptx.common.delete]',
		);
		press('cancel');
		expect(await answer).toBe(false);
	});

	it('confirms, prompts and alerts with the shared labels', async () => {
		const ok = confirmActionDialog('Clear?', context);
		expect(dialog().shadowRoot?.querySelector('button.confirm')?.textContent).toBe(
			'[pptx.common.ok]',
		);
		press('confirm');
		expect(await ok).toBe(true);

		const name = promptNameDialog('New name', 'Old', context);
		const input = dialog().shadowRoot?.querySelector('input') as HTMLInputElement;
		expect(input.value).toBe('Old');
		input.value = 'Fresh';
		press('confirm');
		expect(await name).toBe('Fresh');

		const notice = alertNoticeDialog('Failed', context);
		expect(dialog().shadowRoot?.querySelector('button.cancel')).toBeNull();
		press('confirm');
		await notice;
	});
});
