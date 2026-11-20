// @vitest-environment happy-dom
/**
 * Escape closes the quick-access row's drop-down menus (Presentation options, More actions) and
 * gives focus back to the button that opened them, through the shared `listenForPopupMenuDismiss`.
 */
import React, { act, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { OverflowMenu } from './OverflowMenu';
import { PresentDropdown } from './PresentDropdown';

vi.mock(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

globalThis.IS_REACT_ACT_ENVIRONMENT = true;
const container = document.createElement('div');
document.body.append(container);
const root = createRoot(container);
afterEach(() => act(() => root.render(null)));

function Overflow(): React.ReactElement {
	const [open, setOpen] = useState(false);
	return (
		<OverflowMenu
			isOverflowMenuOpen={open}
			onSetOverflowMenuOpen={setOpen}
			onExportPng={() => {}}
			onPrint={() => {}}
		/>
	);
}

describe('ribbon drop-down menus', () => {
	it.each([
		[
			'pptx.present.optionsTooltip',
			<PresentDropdown key='p' isActive={false} onPresent={() => {}} />,
		],
		['pptx.ribbon.moreActions', <Overflow key='o' />],
	])('closes %s on Escape and gives focus back to its button', (label, element) => {
		act(() => root.render(element));
		const trigger = container.querySelector<HTMLButtonElement>(`[aria-label="${label}"]`)!;
		act(() => trigger.click());
		expect(trigger.getAttribute('aria-expanded')).toBe('true');
		const escape = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
		act(() => {
			document.body.dispatchEvent(escape);
		});
		expect(escape.defaultPrevented).toBe(true);
		expect(trigger.getAttribute('aria-expanded')).toBe('false');
		expect(document.activeElement).toBe(trigger);
	});
});
