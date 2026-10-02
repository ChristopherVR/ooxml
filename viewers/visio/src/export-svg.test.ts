import { afterEach, describe, expect, it, vi } from 'vitest';
import { exportPageSvg, MAX_SVG_EXPORT_BYTES } from './export-svg.js';
import { demoDocument } from './demo-document.js';
import { mountViewer } from './binding.js';
import { RenderResources } from './render-resources.js';
import { renderPage } from './render-svg.js';
import { rasterFixture } from '../tests/raster-fixtures.mjs';

const scene = () => structuredClone(demoDocument);
const parse = (svg: string) => new DOMParser().parseFromString(svg, 'image/svg+xml');
function imageScene(length = 0) {
	const model = scene();
	model.pages[0]!.shapes[0]!.image = {
		...rasterFixture('image/png', length),
		x: 0.1,
		y: 0.2,
		width: 0.5,
		height: 0.7,
	};
	return model;
}
afterEach(() => {
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
	document.body.replaceChildren();
});

describe('portable current-page SVG export', () => {
	it('retains dimensions matching serialized SVG when a serializer callback changes the source', () => {
		const model = scene(),
			serialize = XMLSerializer.prototype.serializeToString;
		const width = model.pages[0]!.width,
			height = model.pages[0]!.height;
		vi.spyOn(XMLSerializer.prototype, 'serializeToString').mockImplementationOnce((node) => {
			const xml = serialize.call(new XMLSerializer(), node);
			model.pages[0]!.width = model.pages[0]!.height = 5000;
			return xml;
		});
		const result = exportPageSvg(model);
		const xml = parse(result.svg);
		expect(result.width).toBe(width);
		expect(result.height).toBe(height);
		expect(xml.documentElement.getAttribute('width')).toBe(`${width}in`);
		expect(xml.documentElement.getAttribute('height')).toBe(`${height}in`);
		expect(JSON.parse(xml.querySelector('metadata')!.textContent!).page).toMatchObject({
			width,
			height,
		});
	});
	it('includes dimensions, accessible title, UTF-8 length and honest compatibility metadata', () => {
		const result = exportPageSvg(demoDocument);
		const xml = parse(result.svg),
			svg = xml.documentElement;
		expect(xml.querySelector('parsererror')).toBeNull();
		expect(svg.getAttribute('width')).toBe('8.5in');
		expect(svg.getAttribute('height')).toBe('7in');
		expect(svg.getAttribute('viewBox')).toBe('0 0 8.5 7');
		expect(svg.getAttribute('role')).toBe('img');
		expect(svg.querySelector('title')?.textContent).toBe('Release workflow');
		expect(svg.querySelector('desc')?.textContent).toContain('not an editable VSDX round-trip');
		expect(result.byteLength).toBe(new TextEncoder().encode(result.svg).byteLength);
		expect(result.diagnostics.some((note) => note.message.includes('Text metrics'))).toBe(true);
		expect(JSON.parse(svg.querySelector('metadata')!.textContent!).diagnostics).toEqual(
			result.diagnostics,
		);
		expect(result).toMatchObject({
			pageIndex: 0,
			pageName: 'Release workflow',
			width: 8.5,
			height: 7,
		});
		expect(result.svg.match(/xmlns="http:\/\/www.w3.org\/2000\/svg"/g)).toHaveLength(1);
	});
	it('preserves valid Unicode, escapes malicious labels and does not create links or executable markup', () => {
		const model = scene(),
			page = model.pages[0]!,
			shape = page.shapes[0]!;
		page.name = '図面 😀 café العربية <script src="https://bad.test/x"/>';
		shape.name = '<img src=x onerror="alert(1)">';
		shape.text.plainText = '😀 <script>alert(1)</script> & "quote" العربية';
		shape.text.fontFamily = 'Arial" onload="alert(1)';
		shape.style.fill = 'url(https://bad.test/paint.svg)';
		const before = structuredClone(model);
		const result = exportPageSvg(model),
			xml = parse(result.svg);
		expect(xml.querySelector('parsererror')).toBeNull();
		expect(xml.querySelector('script, foreignObject, a, img')).toBeNull();
		expect(xml.querySelector('[onload], [onerror], [onclick]')).toBeNull();
		expect(xml.querySelector('title')?.textContent).toBe(page.name);
		expect(xml.querySelector('text')?.textContent).toContain('😀');
		expect(xml.documentElement.textContent).toContain('<script>alert(1)</script>');
		expect(result.svg).toContain('&lt;script');
		expect(
			Array.from(xml.querySelectorAll('image, use')).every((node) =>
				node.getAttributeNS('http://www.w3.org/1999/xlink', 'href')!.startsWith('#'),
			),
		).toBe(true);
		expect(model).toEqual(before);
	});
	it('replaces only XML-invalid controls and isolated surrogates with diagnostics', () => {
		const model = scene();
		model.pages[0]!.name = 'good 😀\u0000\ud800\uffff';
		model.pages[0]!.shapes[0]!.text.plainText = 'x\u000By\udfff';
		model.diagnostics.push({ code: 'test', severity: 'warning', message: 'bad\u0001note' });
		const result = exportPageSvg(model),
			xml = parse(result.svg);
		expect(xml.querySelector('parsererror')).toBeNull();
		expect(result.pageName).toBe('good 😀���');
		expect(xml.querySelector('text')?.textContent).toBe('x�y�');
		expect(result.diagnostics.some((note) => note.message.includes('Invalid XML'))).toBe(true);
		expect(result.diagnostics.some((note) => note.message === 'bad�note')).toBe(true);
	});
	it('omits all viewer selection hooks, roles and tab stops without changing interactive rendering', () => {
		const xml = parse(exportPageSvg(demoDocument).svg);
		expect(
			xml.querySelector(
				'[tabindex], [data-shape-id], [data-page-id], [data-selected], [data-geometry], [role="button"], [aria-pressed], .paper',
			),
		).toBeNull();
		const interactive = renderPage(demoDocument, demoDocument.pages[0]!);
		expect(interactive.svg.querySelector('[data-shape-id][tabindex="0"]')).not.toBeNull();
		interactive.dispose();
	});
	it('merges parser diagnostics and renderer notes without mutating the source diagnostics', () => {
		const model = scene();
		model.diagnostics.push(
			{ code: 'cached', severity: 'warning', message: 'Cached formula approximation.' },
			{ code: 'cached', severity: 'info', message: 'Cached formula approximation.' },
		);
		const result = exportPageSvg(model);
		expect(result.diagnostics).toContainEqual({
			message: 'Cached formula approximation.',
			count: 2,
			severity: 'warning',
		});
		expect(model.diagnostics).toHaveLength(2);
	});
	it('draws only the requested page and its backgrounds in the shared renderer order', () => {
		const model = scene();
		model.pages[0]!.backgroundPageId = model.pages[1]!.id;
		const result = exportPageSvg(model, 0),
			xml = parse(result.svg);
		const titles = Array.from(xml.querySelectorAll('g > title'), (node) => node.textContent);
		expect(titles.indexOf('Your framework')).toBeLessThan(titles.indexOf('Start with an idea'));
		const other = exportPageSvg(model, 1);
		expect(other.pageName).toBe('Architecture');
		expect(parse(other.svg).documentElement.textContent).not.toContain('Start with an idea');
	});
	it.each([-1, 2, 0.1, NaN, Infinity])('rejects invalid page index %s', (index) => {
		expect(() => exportPageSvg(demoDocument, index)).toThrow('page index');
	});
	it('rejects an empty page collection', () => {
		expect(() => exportPageSvg({ format: 'vsdx', pages: [], diagnostics: [] })).toThrow(
			'page index',
		);
	});
});

describe('embedded raster resources and export bounds', () => {
	it.each(['path', 'plainText', 'runText'] as const)(
		'rejects array-valued %s before implicit DOM string expansion',
		(field) => {
			const model = scene(),
				shape = model.pages[0]!.shapes[0]!;
			const value = ['x'.repeat(100_000)];
			if (field === 'path') Object.assign(shape.geometry[0]!, { path: value });
			else if (field === 'plainText') Object.assign(shape.text, { plainText: value });
			else
				shape.text.runs = [
					Object.assign(
						{
							text: 'text',
							fontFamily: 'Arial',
							fontSize: 0.1,
							color: '#000',
							bold: false,
							italic: false,
							underline: false,
						},
						{ text: value },
					),
				] as never;
			const create = vi.spyOn(document, 'createElementNS');
			expect(() => exportPageSvg(model)).toThrow(/invalid (path|plain|run) text/);
			expect(create).not.toHaveBeenCalled();
		},
	);
	it('embeds shared raster bytes once, keeps crop coordinates and releases transient resources', () => {
		const model = imageScene(24_577),
			original = model.pages[0]!.shapes[0]!;
		model.pages[0]!.shapes.push({
			...original,
			id: 'copy',
			image: { ...original.image!, width: 2, height: 3 },
		});
		const create = vi.fn(),
			revoke = vi.fn(),
			dispose = vi.spyOn(RenderResources.prototype, 'dispose');
		vi.stubGlobal(
			'URL',
			Object.assign(class extends URL {}, { createObjectURL: create, revokeObjectURL: revoke }),
		);
		const result = exportPageSvg(model),
			xml = parse(result.svg);
		expect(xml.querySelectorAll('symbol image')).toHaveLength(1);
		expect(xml.querySelectorAll('use')).toHaveLength(2);
		const data = xml
			.querySelector('image')!
			.getAttributeNS('http://www.w3.org/1999/xlink', 'href')!;
		expect(data).toMatch(/^data:image\/png;base64,/);
		expect(Uint8Array.from(atob(data.split(',')[1]!), (point) => point.charCodeAt(0))).toEqual(
			original.image!.bytes,
		);
		expect(
			Array.from(xml.querySelectorAll('use'), (node) =>
				node.getAttributeNS('http://www.w3.org/1999/xlink', 'href'),
			),
		).toEqual(['#visio-export-image-1', '#visio-export-image-1']);
		expect(xml.querySelector('use')?.getAttribute('transform')).toBe(
			'translate(0.1 0.8999999999999999) scale(1 -1)',
		);
		expect(xml.querySelector('symbol')?.getAttribute('preserveAspectRatio')).toBe('none');
		expect(xml.querySelectorAll('clipPath')).toHaveLength(2);
		expect(create).not.toHaveBeenCalled();
		expect(revoke).not.toHaveBeenCalled();
		expect(dispose).toHaveBeenCalledOnce();
		expect(result.svg).not.toContain('blob:');
	});
	it.each(['image/png', 'image/jpeg', 'image/gif'] as const)(
		'embeds only supported raster MIME %s',
		(mimeType) => {
			const model = imageScene();
			Object.assign(model.pages[0]!.shapes[0]!.image!, rasterFixture(mimeType));
			expect(exportPageSvg(model).svg).toContain(`data:${mimeType};base64,`);
		},
	);
	it('rejects executable foreign image MIME before creating DOM', () => {
		const model = imageScene();
		Object.assign(model.pages[0]!.shapes[0]!.image!, { mimeType: 'image/svg+xml' });
		const create = vi.spyOn(document, 'createElementNS');
		expect(() => exportPageSvg(model)).toThrow('Unsupported embedded image');
		expect(create).not.toHaveBeenCalled();
	});
	it('rejects active or corrupt bytes disguised as an allowed raster MIME before DOM allocation', () => {
		const model = imageScene();
		model.pages[0]!.shapes[0]!.image!.bytes = Uint8Array.from(
			new TextEncoder().encode('<svg onload="alert(1)"/>'),
		);
		const create = vi.spyOn(document, 'createElementNS');
		expect(() => exportPageSvg(model)).toThrow('not a supported');
		expect(create).not.toHaveBeenCalled();
	});
	it.each([0, -1, 0.5, NaN, Infinity, MAX_SVG_EXPORT_BYTES + 1])(
		'rejects invalid ceiling %s',
		(maxBytes) => {
			expect(() => exportPageSvg(demoDocument, 0, { maxBytes })).toThrow('maxBytes');
		},
	);
	it('preflights escaped path amplification before renderer allocation', () => {
		const model = scene();
		model.pages[0]!.shapes[0]!.geometry[0]!.path = '&'.repeat(3_000_000);
		const create = vi.spyOn(document, 'createElementNS');
		expect(() => exportPageSvg(model)).toThrow('amplification');
		expect(create).not.toHaveBeenCalled();
	});
	it.each(['startArrow', 'endArrow', 'linePattern', 'startArrowSize', 'endArrowSize'] as const)(
		'rejects oversized nonnumeric %s before allocating DOM or diagnostics',
		(field) => {
			const model = scene();
			Object.assign(model.pages[0]!.shapes[0]!.style, { [field]: 'x'.repeat(100_000) });
			const create = vi.spyOn(document, 'createElementNS');
			expect(() => exportPageSvg(model)).toThrow('invalid');
			expect(create).not.toHaveBeenCalled();
		},
	);
	it('rejects oversized paragraph direction before copying it into SVG attributes', () => {
		const model = scene(),
			text = model.pages[0]!.shapes[0]!.text;
		text.paragraphs = [
			{
				start: 0,
				end: text.plainText.length,
				horizontalAlign: 'left',
				indentLeft: 0,
				indentRight: 0,
				indentFirst: 0,
				spaceBefore: 0,
				spaceAfter: 0,
				lineSpacing: { kind: 'multiple', value: 1.2 },
				direction: 'ltr',
			},
		];
		Object.assign(text.paragraphs[0]!, { direction: 'x'.repeat(100_000) });
		const create = vi.spyOn(document, 'createElementNS');
		expect(() => exportPageSvg(model)).toThrow('paragraph direction');
		expect(create).not.toHaveBeenCalled();
	});
	it('rejects invalid diagnostic severity before serializing metadata', () => {
		const model = scene();
		model.diagnostics = [{ code: 'host', severity: 'warning', message: 'Cached value' }];
		Object.assign(model.diagnostics[0]!, { severity: 'x'.repeat(100_000) });
		const create = vi.spyOn(document, 'createElementNS');
		expect(() => exportPageSvg(model)).toThrow('diagnostic severity');
		expect(create).not.toHaveBeenCalled();
	});
	it('preflights repeated text-style amplification', () => {
		const model = scene(),
			text = model.pages[0]!.shapes[0]!.text;
		text.plainText = 'x'.repeat(4000);
		text.fontFamily = 'f'.repeat(1024);
		const create = vi.spyOn(document, 'createElementNS');
		expect(() => exportPageSvg(model)).toThrow('amplification');
		expect(create).not.toHaveBeenCalled();
	});
	it('preflights base64 expansion before encoding or creating image nodes', () => {
		const model = imageScene(1_000_000);
		const encode = vi.spyOn(globalThis, 'btoa'),
			create = vi.spyOn(document, 'createElementNS');
		expect(() => exportPageSvg(model, 0, { maxBytes: 1_000_000 })).toThrow('amplification');
		expect(encode).not.toHaveBeenCalled();
		expect(create).not.toHaveBeenCalled();
	});
	it('cleans resources when serialization throws', () => {
		const dispose = vi.spyOn(RenderResources.prototype, 'dispose');
		vi.spyOn(XMLSerializer.prototype, 'serializeToString').mockImplementation(() => {
			throw new Error('serialization failed');
		});
		expect(() => exportPageSvg(imageScene())).toThrow('serialization failed');
		expect(dispose).toHaveBeenCalledOnce();
	});
	it('cleans partially rendered resources when encoding fails', () => {
		const dispose = vi.spyOn(RenderResources.prototype, 'dispose');
		vi.spyOn(globalThis, 'btoa').mockImplementation(() => {
			throw new Error('encoding failed');
		});
		expect(() => exportPageSvg(imageScene())).toThrow('encoding failed');
		expect(dispose).toHaveBeenCalledOnce();
	});
	it('checks the exact final UTF-8 length and cleans resources on failure', () => {
		const dispose = vi.spyOn(RenderResources.prototype, 'dispose');
		vi.spyOn(XMLSerializer.prototype, 'serializeToString').mockReturnValue('😀'.repeat(250_001));
		expect(() => exportPageSvg(demoDocument, 0, { maxBytes: 1_000_000 })).toThrow('amplification');
		expect(dispose).toHaveBeenCalledOnce();
	});
	it('rejects a missing DOM explicitly at invocation without affecting imports', () => {
		vi.stubGlobal('XMLSerializer', undefined);
		expect(() => exportPageSvg(demoDocument)).toThrow('requires a browser DOM');
	});
});

describe('shared export method lifecycle', () => {
	it('exports the current page without changing document, zoom, selection or events', () => {
		const events = vi.fn();
		const viewer = mountViewer(document.createElement('div'), {
			document: demoDocument,
			pageIndex: 1,
			zoom: 3,
		});
		viewer.controller.selectShape({ id: 'a1', name: 'Your framework', pageId: '2' });
		viewer.controller.onEvent(events);
		const state = viewer.controller.state;
		const before = viewer.element.shadowRoot!.innerHTML;
		const result = viewer.exportSvg();
		expect(result.pageIndex).toBe(1);
		expect(viewer.controller.state).toBe(state);
		expect(viewer.element.shadowRoot!.innerHTML).toBe(before);
		expect(events).not.toHaveBeenCalled();
		expect(viewer.element.exportSvg().pageIndex).toBe(1);
		viewer.destroy();
		expect(() => viewer.exportSvg()).toThrow('destroyed');
		expect(() => viewer.element.exportSvg()).toThrow('destroyed');
	});
	it('rejects no document and retains selection/state when export fails', () => {
		const viewer = mountViewer(document.createElement('div'));
		expect(() => viewer.exportSvg()).toThrow('Open a document');
		viewer.update({ document: demoDocument });
		viewer.controller.selectShape({ id: 's1', name: 'Start with an idea' });
		const state = viewer.controller.state;
		expect(() => viewer.exportSvg({ maxBytes: 10 })).toThrow('amplification');
		expect(viewer.controller.state).toBe(state);
		viewer.destroy();
	});
});
