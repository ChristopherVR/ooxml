/**
 * The thumbnail rail's field and table style contexts: built with the shared
 * slides-pane helpers from the rail's inputs, each falling back to the deck
 * loader so an editor-bound rail matches the canvas.
 */
import { signal } from '@angular/core';
import { buildSlidesPaneFieldContext } from 'ooxml-ui/pptx';
import type { ParsedTableStyleMap, PptxSlide, PptxTheme } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import type { LoadContentService } from './load-content.service';
import { createSlidesPanelContexts } from './slides-panel-contexts';
import { tableStyleContextFor } from './table-renderer-helpers';

const THEME = { colorScheme: { accent1: '#112233' } } as unknown as PptxTheme;
const STYLES = { styles: {} } as unknown as ParsedTableStyleMap;

function inputs(overrides: Partial<Record<string, unknown>> = {}) {
	return {
		headerFooter: signal(overrides['headerFooter'] as never),
		customProperties: signal(overrides['customProperties'] as never),
		firstSlideNumber: signal(overrides['firstSlideNumber'] as number | undefined),
		locale: signal(overrides['locale'] as string | undefined),
		theme: signal(overrides['theme'] as PptxTheme | undefined),
		tableStyleMap: signal(overrides['tableStyleMap'] as ParsedTableStyleMap | undefined),
	};
}

const loader = {
	headerFooter: signal({ footerText: 'Loader footer' }),
	customProperties: signal([]),
	firstSlideNumber: signal(7),
	theme: signal(THEME),
	tableStyleMap: signal(STYLES),
} as unknown as LoadContentService;

describe('slides panel thumbnail contexts', () => {
	it('builds the deck fields from the inputs with the shared helper', () => {
		const contexts = createSlidesPanelContexts(
			inputs({ firstSlideNumber: 45, headerFooter: { footerText: 'Host' }, locale: 'de-DE' }),
			null,
		);
		expect(contexts.fields.deckContext()).toStrictEqual(
			buildSlidesPaneFieldContext({
				headerFooter: { footerText: 'Host' },
				firstSlideNumber: 45,
				locale: 'de-DE',
			}),
		);
		const slide = { id: 's', slideNumber: 2, elements: [] } as unknown as PptxSlide;
		expect(contexts.fields.forSlide(slide)).toMatchObject({ slideNumber: 2, firstSlideNumber: 45 });
	});

	it('falls back to the deck loader per source', () => {
		const contexts = createSlidesPanelContexts(inputs({ firstSlideNumber: 45 }), loader);
		expect(contexts.fields.deckContext()).toMatchObject({
			footerText: 'Loader footer',
			firstSlideNumber: 45,
		});
		expect(contexts.tableStyle()).toStrictEqual(tableStyleContextFor(STYLES, THEME));
	});

	it('prefers the host theme and has no table context without theme or styles', () => {
		const other = { colorScheme: { accent1: '#445566' } } as unknown as PptxTheme;
		expect(createSlidesPanelContexts(inputs({ theme: other }), loader).tableStyle()).toStrictEqual(
			tableStyleContextFor(STYLES, other),
		);
		expect(createSlidesPanelContexts(inputs(), null).tableStyle()).toBeUndefined();
	});
});
