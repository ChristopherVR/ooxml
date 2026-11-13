import type { PptxCustomShow, PptxPresentationProperties } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { allButtons } from './dialog-footer.test-support';
import { openSlideShowDialog } from './slide-show-dialog';

const SHOWS: PptxCustomShow[] = [{ id: '4', name: 'Short', slideRIds: ['rId1'] }];

function openAndSave(act: () => void): PptxPresentationProperties {
	document.body.replaceChildren();
	const onSave = vi.fn();
	openSlideShowDialog(document, createTranslator(), {}, 6, onSave, SHOWS);
	act();
	allButtons(document)
		.find((button) => button.textContent?.trim() === 'OK')
		?.click();
	expect(onSave).toHaveBeenCalledTimes(1);
	return onSave.mock.calls[0]?.[0] as PptxPresentationProperties;
}

// Order: show type x3, all slides, from/to, custom show, advance x2.
const radio = (index: number) =>
	document.querySelectorAll<HTMLInputElement>('pptx-ui-radio')[index]!;
const change = (input: HTMLInputElement) => {
	input.checked = true;
	input.dispatchEvent(new Event('change', { bubbles: true }));
};

describe('vanilla set up slide show dialog follows the shared rules', () => {
	it('a kiosk show loops continuously', () => {
		const saved = openAndSave(() => change(radio(2)));
		expect(saved).toMatchObject({ showType: 'kiosk', loopContinuously: true });
	});

	it('a custom show defaults to the first show', () => {
		const saved = openAndSave(() => {
			const custom = document.querySelector<HTMLInputElement>('[data-pptx-show-slides-custom]')!;
			change(custom);
		});
		expect(saved).toMatchObject({ showSlidesMode: 'customShow', showSlidesCustomShowId: '4' });
	});

	it('a range starts as the whole deck', () => {
		const saved = openAndSave(() => change(radio(4)));
		expect(saved).toMatchObject({ showSlidesMode: 'range', showSlidesFrom: 1, showSlidesTo: 6 });
	});
});
