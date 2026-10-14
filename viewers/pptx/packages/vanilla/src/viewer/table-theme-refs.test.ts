// @vitest-environment happy-dom
/**
 * "Themed Style 1" draws its background and gridlines through the theme's
 * format scheme (`a:tblBg/a:fillRef`, `a:lnRef`). The loader must keep the
 * deck's format scheme in state and the stage must hand it to the table
 * renderer, or such a table renders with no background and no lines.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { PptxHandler } from 'pptx-viewer-core';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { createLoadingController } from './loading-controller';
import { createDefaultRegistry } from './render';
import { createRenderController } from './render-controller';
import type { RenderControllerDeps } from './render-controller';
import { createInitialViewerState, createStore } from './state';

const THEMED_1_ACCENT_1 = '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}';
const FIXTURE = resolve(process.cwd(), '../../../../e2e/pptx/fixtures/table-styling.pptx');

/** The fixture with its first table switched to Themed Style 1 and every option off. */
async function themedDeck(): Promise<ArrayBuffer> {
	const handler = new PptxHandler();
	const buf = readFileSync(FIXTURE);
	const deck = await handler.load(
		buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
	);
	const [slide] = deck.slides;
	const elements = slide.elements.map((element) =>
		element.type === 'table' && element.tableData
			? ({
					...element,
					tableData: {
						...element.tableData,
						tableStyleId: THEMED_1_ACCENT_1,
						firstRowHeader: false,
						bandedRows: false,
					},
				} as PptxElement)
			: element,
	);
	const bytes = await handler.save([{ ...slide, elements }]);
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

describe('vanilla table rendering with theme style references', () => {
	it('keeps the format scheme and paints the Themed Style table background', async () => {
		const store = createStore(createInitialViewerState());
		const loader = createLoadingController({
			options: {},
			store,
			getTranslator: () => (key) => key,
			getEditor: () => undefined,
		});
		await loader.load(await themedDeck());
		try {
			const state = store.get();
			expect(state.formatScheme?.lineStyles.length).toBeGreaterThan(0);
			const deps = {
				doc: document,
				store,
				registry: createDefaultRegistry(),
				getTranslator: () => (key: string) => key,
			};
			// Every other dependency is a getter or callback this render does not need.
			const renderer = createRenderController(
				new Proxy(deps, {
					get: (target, key) =>
						key in target ? target[key as keyof typeof target] : () => undefined,
				}) as unknown as RenderControllerDeps,
			);
			const node = renderer.renderSlideNode(state.slides[0], 1);
			const cell = node.querySelector<HTMLElement>('td');
			// The table background is the theme's second fill style (`a:tblBg/
			// a:fillRef idx="2"`), a gradient of accent1 in this deck's theme. The
			// gridlines resolve too, but happy-dom drops their `max()` widths.
			expect(cell?.style.background).toContain('linear-gradient');
		} finally {
			loader.releaseLoaded();
		}
	});
});
