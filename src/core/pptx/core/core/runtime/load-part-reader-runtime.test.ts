import { readFileSync } from 'node:fs';

import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { XmlObject } from '../../types';
import { toStrictNamespaceUri } from '../../utils/strict-namespace-map';
import type { LoadPartReader } from './load-part-reader';
import { PptxHandlerRuntime } from './PptxHandlerRuntimeImplementation';
import { PptxHandlerRuntime as RuntimeState } from './PptxHandlerRuntimeState';

const relsPath = 'ppt/slides/_rels/slide1.xml.rels';

function deck(path: string): ArrayBuffer {
	const bytes = readFileSync(new URL(`../../../__tests__/fixtures/e2e/${path}`, import.meta.url));
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

class StateRuntime extends RuntimeState {
	constructor() {
		super();
		this.zip = new JSZip().file(
			relsPath,
			'<Relationships><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout" Target="layout.xml"/></Relationships>',
		);
		this.parser = new XMLParser({ ignoreAttributes: false });
	}
	get reader() {
		return this.loadPartReader;
	}
	begin() {
		return this.beginPartReads();
	}
	end(reader?: LoadPartReader) {
		this.endPartReads(reader);
	}
	read() {
		return this.readPartXml(relsPath);
	}
	normalize() {
		this.detectAndSetStrictConformance({
			'p:presentation': { '@_xmlns:p': 'http://purl.oclc.org/ooxml/presentationml/main' },
		});
	}
}

/** Loads exactly as before the cache: every reader closes as soon as it opens. */
class UncachedRuntime extends PptxHandlerRuntime {
	protected override beginPartReads(): LoadPartReader {
		const reader = super.beginPartReads();
		this.endPartReads(reader);
		return reader;
	}
}

/** Counts every inflation of every package entry during a load. */
class ObservedRuntime extends PptxHandlerRuntime {
	inflations = new Map<string, number>();
	failBuild = false;
	get reader() {
		return this.loadPartReader;
	}
	protected override beginPartReads(): LoadPartReader {
		this.inflations.clear();
		for (const file of Object.values(this.zip.files)) {
			if (!file.dir) {
				const inflate = file.async.bind(file) as (...args: unknown[]) => Promise<unknown>;
				vi.spyOn(file, 'async').mockImplementation(((...args: unknown[]) => {
					this.inflations.set(file.name, (this.inflations.get(file.name) ?? 0) + 1);
					return inflate(...args);
				}) as typeof file.async);
			}
		}
		return super.beginPartReads();
	}
	protected override async buildLoadData(...args: Parameters<PptxHandlerRuntime['buildLoadData']>) {
		if (this.failBuild) {
			throw new Error('build failure');
		}
		return super.buildLoadData(...args);
	}
	repeated(): string[] {
		return [...this.inflations].filter(([, count]) => count > 1).map(([name]) => name);
	}
}

function relationship(data: XmlObject | undefined): XmlObject {
	return (data!.Relationships as XmlObject).Relationship as XmlObject;
}

afterEach(() => {
	vi.restoreAllMocks();
});

describe('runtime part reads during load', () => {
	it.each(['sample-deck.pptx', 'three-d-parity/three-d-charts.pptx'])(
		'inflates each package part of %s at most once and parses it identically',
		async (path) => {
			const observed = new ObservedRuntime();
			const control = new UncachedRuntime();
			try {
				const cached = await observed.load(deck(path));
				expect(observed.inflations.size).toBeGreaterThan(10);
				expect(observed.repeated()).toEqual([]);
				expect(JSON.stringify(cached)).toBe(JSON.stringify(await control.load(deck(path))));
			} finally {
				observed.dispose();
				control.dispose();
			}
		},
	);

	it('releases the cache when a load settles, so save and the next load read the archive', async () => {
		const runtime = new ObservedRuntime();
		const first = await runtime.load(deck('sample-deck.pptx'));
		expect(runtime.reader).toBeUndefined();
		expect((await runtime.save(first.slides)).byteLength).toBeGreaterThan(0);
		expect(runtime.reader).toBeUndefined();
		const second = await runtime.load(deck('sample-deck.pptx'));
		expect(JSON.stringify(second)).toBe(JSON.stringify(first));
		expect(runtime.repeated()).toEqual([]);
		runtime.dispose();
	});

	it('releases the cache after a late load failure and can load again', async () => {
		const runtime = new ObservedRuntime();
		runtime.failBuild = true;
		await expect(runtime.load(deck('sample-deck.pptx'))).rejects.toThrow('build failure');
		expect(runtime.reader).toBeUndefined();
		runtime.failBuild = false;
		expect((await runtime.load(deck('sample-deck.pptx'))).slides.length).toBeGreaterThan(0);
		runtime.dispose();
	});

	it('keeps concurrent loads in separate handlers apart', async () => {
		const a = new PptxHandlerRuntime();
		const b = new PptxHandlerRuntime();
		const control = new UncachedRuntime();
		try {
			const [first, second] = await Promise.all([
				a.load(deck('sample-deck.pptx')),
				b.load(deck('three-d-parity/three-d-charts.pptx')),
			]);
			expect(JSON.stringify(first)).toBe(
				JSON.stringify(await control.load(deck('sample-deck.pptx'))),
			);
			expect(JSON.stringify(second)).toBe(
				JSON.stringify(await control.load(deck('three-d-parity/three-d-charts.pptx'))),
			);
		} finally {
			a.dispose();
			b.dispose();
			control.dispose();
		}
	});

	it('loads and saves a Strict package identically with and without the cache', async () => {
		const zip = await JSZip.loadAsync(deck('sample-deck.pptx'));
		for (const file of Object.values(zip.files)) {
			if (!file.dir && /\.(xml|rels)$/u.test(file.name)) {
				const xml = await file.async('string');
				zip.file(
					file.name,
					xml.replace(/http:\/\/schemas\.openxmlformats\.org\/[^\s"'<>]+/gu, toStrictNamespaceUri),
				);
			}
		}
		const bytes = await zip.generateAsync({ type: 'arraybuffer' });
		const cached = new PptxHandlerRuntime();
		const uncached = new UncachedRuntime();
		try {
			const first = await cached.load(bytes);
			const second = await uncached.load(bytes);
			expect(first.conformance).toBe('strict');
			expect(JSON.stringify(first)).toBe(JSON.stringify(second));
			const saved = await JSZip.loadAsync(await cached.save(first.slides));
			const control = await JSZip.loadAsync(await uncached.save(second.slides));
			expect(Object.keys(saved.files).sort()).toEqual(Object.keys(control.files).sort());
			const omitTimes = (xml: string) =>
				xml.replace(
					/(<dcterms:(?:created|modified)\b[^>]*>)[^<]*(<\/dcterms:(?:created|modified)>)/gu,
					'$1[TIMESTAMP]$2',
				);
			for (const file of Object.values(saved.files)) {
				if (file.dir) {
					continue;
				}
				if (file.name === 'docProps/core.xml') {
					expect(omitTimes(await file.async('string'))).toBe(
						omitTimes(await control.file(file.name)!.async('string')),
					);
				} else {
					expect(await file.async('uint8array')).toEqual(
						await control.file(file.name)!.async('uint8array'),
					);
				}
			}
		} finally {
			cached.dispose();
			uncached.dispose();
		}
	});
});

describe('runtime part reader lifecycle', () => {
	it('normalizes Strict relationships even when parsed before Strict detection', async () => {
		const runtime = new StateRuntime();
		runtime.begin();
		expect(relationship(await runtime.read())['@_Type']).toContain('purl.oclc.org');
		runtime.normalize();
		expect(relationship(await runtime.read())['@_Type']).toBe(
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
		);
		runtime.end();
	});

	it('closes only the reader a load opened, never a newer one', () => {
		const runtime = new StateRuntime();
		const older = runtime.begin();
		const newer = runtime.begin();
		expect(older.isOpen).toBe(false);
		runtime.end(older);
		expect(newer.isOpen).toBe(true);
		expect(runtime.reader).toBe(newer);
		runtime.end(newer);
		expect(runtime.reader).toBeUndefined();
	});

	it('releases an open reader on dispose', () => {
		const runtime = new (class extends PptxHandlerRuntime {
			open() {
				this.zip = new JSZip();
				return this.beginPartReads();
			}
			get reader() {
				return this.loadPartReader;
			}
		})();
		const reader = runtime.open();
		runtime.dispose();
		expect(reader.isOpen).toBe(false);
		expect(runtime.reader).toBeUndefined();
	});
});
