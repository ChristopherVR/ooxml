/**
 * The element render error contract every binding's building blocks pass (React, Vue, Svelte,
 * Angular, vanilla), beside `building-blocks-contract.mjs`: a slide element whose renderer
 * throws is omitted on its own (canvas and slides pane), the host's `onElementRenderError`
 * hears about it with the shared payload, and a reporter that throws, or none at all, falls back
 * to `console.error` without taking the blocks down. The fault is injected through the shared
 * test hook `probeElementRenderFault` (`Symbol.for(ELEMENT_RENDER_FAULT_KEY)` set to an element
 * id), so the failure comes from inside each binding's own element renderer.
 *
 * Each binding supplies a building-blocks harness (see `building-blocks-contract.d.mts`) whose
 * `mount` forwards every option, `onElementRenderError` included.
 */

const FAULT = Symbol.for('ooxml-ui.pptx.element-render-fault');
const HEALTHY = 'Healthy element';
const BROKEN = 'Broken element';

/** A one-slide deck with a healthy and a broken text box. */
export async function elementRenderErrorContractDeck(Handler) {
	const { handler, data } = await Handler.create({ initialSlideCount: 1 });
	try {
		data.slides[0].elements = [
			{ id: 'healthy', type: 'text', x: 40, y: 40, width: 400, height: 60, text: HEALTHY },
			{ id: 'broken', type: 'text', x: 40, y: 160, width: 400, height: 60, text: BROKEN },
		];
		data.slides[0].isDirty = true;
		return await handler.save(data.slides);
	} finally {
		handler.dispose();
	}
}

async function until(harness, check, label) {
	const deadline = Date.now() + 20000;
	for (;;) {
		await harness.settle();
		let failure;
		try {
			check();
			return;
		} catch (error) {
			failure = error;
		}
		if (Date.now() > deadline) {
			throw new Error(`${harness.name}: timed out waiting for ${label}: ${String(failure)}`);
		}
		await new Promise((resolve) => {
			setTimeout(resolve, 10);
		});
	}
}

/** The id the loaded deck gave the element showing `text` on the canvas. */
function idOf(host, text) {
	const node = [...host.querySelectorAll('[data-element-id]')].find((candidate) =>
		candidate.textContent?.includes(text),
	);
	return node?.getAttribute('data-element-id') ?? null;
}

/**
 * Run the contract against `harness` (a building-blocks harness). `expect` and `vi` are the
 * caller's Vitest.
 */
export async function runElementRenderErrorContract({ harness, expect, vi, Handler }) {
	const content = await elementRenderErrorContractDeck(Handler);
	const base = {
		content,
		canEdit: false,
		googleFonts: false,
		onRequest: () => {},
		onDirtyChange: () => {},
	};
	const mountWith = async (extra) => {
		const host = document.createElement('div');
		document.body.append(host);
		await harness.mount(host, { ...base, ...extra });
		await until(harness, () => expect(host.textContent).toContain(HEALTHY), 'the healthy element');
		return host;
	};
	const unmount = (host) => {
		harness.destroy();
		host.remove();
	};
	const log = vi.spyOn(console, 'error').mockImplementation(() => {});
	try {
		// 1. Without a fault both elements render; learn the id the deck loaded with.
		let host = await mountWith({});
		await until(harness, () => expect(idOf(host, BROKEN)).not.toBeNull(), 'the broken element');
		const brokenId = idOf(host, BROKEN);
		expect(idOf(host, HEALTHY)).not.toBe(brokenId);
		unmount(host);

		// 2. The faulty element alone is omitted, everywhere it is drawn, and reported.
		globalThis[FAULT] = brokenId;
		log.mockClear();
		const report = vi.fn();
		host = await mountWith({ onElementRenderError: report });
		await until(harness, () => expect(report).toHaveBeenCalled(), 'the report');
		expect(host.textContent).not.toContain(BROKEN);
		expect(idOf(host, HEALTHY)).not.toBeNull();
		expect(harness.state().error).toBeNull();
		expect(harness.state().slideCount).toBe(1);
		// With a reporter, no surface (canvas, slides pane) falls back to the console.
		expect(log).not.toHaveBeenCalledWith(
			expect.any(Error),
			expect.objectContaining({ elementId: brokenId }),
		);
		for (const [error, context] of report.mock.calls) {
			expect(error).toBeInstanceOf(Error);
			expect(context).toEqual({
				elementId: brokenId,
				elementType: expect.any(String),
				slideId: expect.any(String),
			});
		}
		unmount(host);

		// 3. A reporter that throws falls back to console.error; the blocks stay.
		log.mockClear();
		const throwing = vi.fn(() => {
			throw new Error('host reporter failed');
		});
		host = await mountWith({ onElementRenderError: throwing });
		await until(harness, () => expect(throwing).toHaveBeenCalled(), 'the throwing reporter');
		expect(log).toHaveBeenCalledWith(
			expect.any(Error),
			expect.objectContaining({ elementId: brokenId }),
		);
		expect(host.textContent).not.toContain(BROKEN);
		expect(harness.state().error).toBeNull();
		unmount(host);

		// 4. Without a reporter the default is console.error(error, context).
		log.mockClear();
		host = await mountWith({});
		await until(
			harness,
			() =>
				expect(log).toHaveBeenCalledWith(
					expect.any(Error),
					expect.objectContaining({ elementId: brokenId }),
				),
			'the default report',
		);
		expect(host.textContent).not.toContain(BROKEN);
		unmount(host);
	} finally {
		Reflect.deleteProperty(globalThis, FAULT);
		log.mockRestore();
	}
}
