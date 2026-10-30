// @vitest-environment happy-dom
/**
 * CustomShowsDialog: New, Copy, Remove and the two slide lists with Add, Remove and Move
 * Up/Down. The decisions live in the shared controller; this pins the React mapping onto it.
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { CustomShowsDialog } from './CustomShowsDialog';

vi.mock<typeof import('../i18n/use-translation')>(import('../i18n/use-translation'), () => ({
	useTranslation: () => ({
		t: (key: string) => translationsEn[key] ?? key,
	}),
}));

const slides = [1, 2, 3].map(
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
const shows: PptxCustomShow[] = [
	{ id: '7', name: 'Gone', slideRIds: [] },
	{ id: '8', name: 'Kept', slideRIds: ['rId1', 'rId2'] },
];

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
});

const click = (element: Element | undefined): void => {
	act(() => {
		(element as HTMLElement).click();
	});
};

describe('customShowsDialog', () => {
	it('copies a show, moves its slides and clears jumps into a removed show', () => {
		const onSave = vi.fn();
		act(() => {
			root.render(
				<CustomShowsDialog
					customShows={shows}
					slides={slides}
					activeCustomShowId='8'
					onSave={onSave}
					onSetActive={vi.fn()}
					onRun={vi.fn()}
					onClose={vi.fn()}
				/>,
			);
		});
		const button = (label: string, scope: ParentNode = container) =>
			Array.from(scope.querySelectorAll('button')).find((b) => b.textContent === label);

		click(button('Copy'));
		click(container.querySelectorAll('[aria-label="Slides in show"] [role="option"]')[1]);
		click(button('Move up'));
		click(container.querySelectorAll('[aria-label="Saved Shows"] [role="option"]')[0]);
		click(button('Remove'));
		const footer = container.querySelector('pptx-ui-dialog-footer')!.shadowRoot!;
		click(button('OK', footer));

		const [saved, cleared] = onSave.mock.calls[0] as [PptxCustomShow[], PptxSlide[] | undefined];
		expect(saved.map((show) => show.name)).toEqual(['Kept', 'Copy of Kept']);
		expect(saved[1]?.slideRIds).toEqual(['rId2', 'rId1']);
		expect(cleared?.[0]?.elements[0]).not.toHaveProperty('actionClick');
	});

	it('restricts playback from the picker and offers every show', () => {
		const onSetActive = vi.fn();
		act(() => {
			root.render(
				<CustomShowsDialog
					customShows={shows}
					slides={slides}
					activeCustomShowId={null}
					onSave={vi.fn()}
					onSetActive={onSetActive}
					onRun={vi.fn()}
					onClose={vi.fn()}
				/>,
			);
		});
		const options = Array.from(container.querySelectorAll('option')).map((o) => o.textContent);
		expect(options).toEqual(['All Slides', 'Gone', 'Kept']);
	});
});
