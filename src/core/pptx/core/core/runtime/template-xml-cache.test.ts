import { readFile } from 'node:fs/promises';
import { expect, it, vi } from 'vitest';

import { PptxHandlerRuntime } from './PptxHandlerRuntimeImplementation';
import type { LoadPartReader } from './load-part-reader';

class ObservedRuntime extends PptxHandlerRuntime {
	readonly templateParses = new Map<string, number>();
	private restoreParser: (() => void) | undefined;

	protected override beginPartReads(): LoadPartReader {
		const parse = this.parser.parse.bind(this.parser);
		const spy = vi.spyOn(this.parser, 'parse').mockImplementation((xml, ...options) => {
			if (typeof xml === 'string' && /<(?:p:sldMaster|p:sldLayout|a:theme)\b/u.test(xml)) {
				this.templateParses.set(xml, (this.templateParses.get(xml) ?? 0) + 1);
			}
			return parse(xml, ...options);
		});
		this.restoreParser = () => spy.mockRestore();
		return super.beginPartReads();
	}

	override dispose(): void {
		this.restoreParser?.();
		super.dispose();
	}
}

it('parses template and theme object trees once while loading a deck', async () => {
	const source = await readFile(
		new URL('../../../__tests__/fixtures/e2e/sample-deck.pptx', import.meta.url),
	);
	const runtime = new ObservedRuntime();
	try {
		const data = await runtime.load(new Uint8Array(source).buffer);
		expect(data.slideMasters?.[0]?.layouts?.length).toBeGreaterThan(0);
		expect(data.themeOptions?.length).toBeGreaterThan(0);
		expect(runtime.templateParses.size).toBeGreaterThan(2);
		expect(Math.max(...runtime.templateParses.values())).toBe(1);
	} finally {
		runtime.dispose();
	}
});
