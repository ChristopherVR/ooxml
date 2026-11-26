import { type Command, type EditorContext, unavailableReason } from 'ooxml-core/xlsx/ui';
import type { MenuEntry } from './popover';

/** Translated reason for a disabled command (the shared core rule). */
export function disabledReason(ctx: EditorContext, command: Command): string | undefined {
	return unavailableReason(ctx, command);
}

export function describeDisabled(button: HTMLButtonElement, title: string, reason?: string): void {
	button.title = reason ? `${title}: ${reason}` : title;
	if (reason) button.setAttribute('aria-description', reason);
	else button.removeAttribute('aria-description');
}

export function updateMenuAvailability(
	button: HTMLButtonElement,
	title: string,
	entries: MenuEntry[],
): void {
	const items = entries.flatMap((entry) => ('separator' in entry ? [] : [entry]));
	button.disabled = items.every(({ disabled }) => disabled);
	const reason = items[0]?.reason;
	describeDisabled(
		button,
		title,
		button.disabled && items.every((item) => item.reason === reason) ? reason : undefined,
	);
}
