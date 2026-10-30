import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { openCustomShowsDialog } from './custom-shows-dialog';
import { allButtons } from './dialog-footer.test-support';

const slides = (): PptxSlide[] =>
	[1, 2, 3].map(
		(n) =>
			({
				id: `s${n}`,
				rId: `rId${n}`,
				slideNumber: n,
				elements:
					n === 1
						? [{ id: 'e1', type: 'shape', actionClick: { action: 'ppaction://customshow?id=7' } }]
						: [],
			}) as unknown as PptxSlide,
	);

const all = <T extends Element>(selector: string): T[] =>
	Array.from(document.querySelectorAll<T>(selector));

describe('vanilla custom-shows dialog editing', () => {
	it('copies a show, reorders its slides and clears jumps into a removed show', () => {
		document.body.replaceChildren();
		const onSave = vi.fn();
		openCustomShowsDialog(document, createTranslator(), {
			shows: [
				{ id: '7', name: 'Gone', slideRIds: [] },
				{ id: '8', name: 'Kept', slideRIds: ['rId1', 'rId2'] },
			],
			slides: slides(),
			activeShowId: '8',
			onSave,
			onSetActive: vi.fn(),
			onRun: vi.fn(),
		});
		const columnButton = (label: string) =>
			all<HTMLButtonElement>('.pptxv-custom-shows-list button').find(
				(b) => b.textContent === label,
			)!;
		const button = (label: string) => allButtons(document).find((b) => b.textContent === label)!;

		columnButton('Copy').click();
		expect(all('[aria-label="Saved Shows"] li')).toHaveLength(3);
		// The copy is selected; select its second slide and move it up.
		all<HTMLElement>('[aria-label="Slides in show"] li')[1]!.click();
		button('Move up').click();
		// Remove the first show ("Gone") from the dialog.
		all<HTMLElement>('[aria-label="Saved Shows"] li')[0]!.click();
		columnButton('Remove').click();
		button('OK').click();

		const [shows, cleared] = onSave.mock.calls[0] as [PptxCustomShow[], PptxSlide[] | undefined];
		expect(shows.map((s) => s.name)).toEqual(['Kept', 'Copy of Kept']);
		expect(shows[1]?.slideRIds).toEqual(['rId2', 'rId1']);
		expect(cleared?.[0]?.elements[0]).not.toHaveProperty('actionClick');
	});
});
