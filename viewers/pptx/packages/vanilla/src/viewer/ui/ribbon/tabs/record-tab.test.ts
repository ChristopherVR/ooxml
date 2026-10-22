import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../../i18n';
import type { RibbonSlideShowHandlers } from '../ribbon-types';
import { createRecordTab } from './record-tab';

describe('record shared commands', () => {
	it('starts the rehearsal from both recording commands and disables the placeholders', () => {
		const startRehearsal = vi.fn();
		const el = createRecordTab(document, createTranslator(), {
			startRehearsal,
		} as unknown as RibbonSlideShowHandlers);
		expect(el.querySelectorAll('pptx-ui-ribbon-section > pptx-ui-ribbon-group')).toHaveLength(4);
		expect(
			[...el.querySelectorAll('pptx-ui-ribbon-command[disabled]')].map((command) =>
				command.getAttribute('data-ribbon-control'),
			),
		).toStrictEqual([
			'record.camera.cameo',
			'record.manage.clear',
			'record.manage.reset',
			'record.help.learnMore',
		]);
		for (const id of ['record.record.fromBeginning', 'record.record.fromCurrent']) {
			el.querySelector(`[data-ribbon-control="${id}"]`)!
				.shadowRoot!.querySelector('button')!
				.click();
		}
		expect(startRehearsal).toHaveBeenCalledTimes(2);
	});
});
