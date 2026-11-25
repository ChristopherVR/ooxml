import { describe, expect, it } from 'vitest';
import { drawSvgSubtree } from './pdf-svg';
import { childElements, parseSvgMarkup } from './pdf-svg-dom';
import { readShadowFilter } from './pdf-svg-shadow';
import { PdfVectorPage } from './pdf-vector';

const SVG = 'xmlns="http://www.w3.org/2000/svg"';

async function draw(markup: string) {
	const page = new PdfVectorPage(400, 400);
	const notes = drawSvgSubtree(page, (await parseSvgMarkup(markup))!, new Map());
	return { page, notes, content: page.content };
}

describe('soft shadows in the SVG converter', () => {
	it('draws a blurred shadow group (Visio) as one soft-masked picture instead of hard paths', async () => {
		const { page, notes, content } = await draw(
			`<svg ${SVG}><defs><filter id="f" filterUnits="userSpaceOnUse"><feGaussianBlur stdDeviation="4"/></filter></defs>` +
				`<g data-shadow="" transform="translate(5 5)" opacity="0.5" filter="url(#f)"><path d="M 10 10 L 90 10 L 90 60 L 10 60 Z" fill="#336699" stroke="#336699"/></g>` +
				`<path d="M 10 10 L 90 10 L 90 60 L 10 60 Z" fill="#ffffff"/></svg>`,
		);
		expect(notes).not.toContain('shadowBlurDropped');
		// The group's offset, then the picture at half opacity, then the shape over it.
		expect(content).toMatch(
			/^q\n1 0 0 1 5 5 cm\nq\n\/GS0 gs\n[^\n]+ cm\n\/Im0 Do\nQ\nQ\nq\n(?:\/GS\d gs\n)?1 1 1 rg\n/,
		);
		expect(content).not.toContain('0.2 0.4 0.6 rg');
		const image = page.images.get('Im0')!;
		expect(image.alpha).toBeDefined();
		expect([...image.data.subarray(0, 3)]).toEqual([51, 102, 153]);
		expect(page.alphas.get('GS0')).toEqual({ fill: 0.5, stroke: 0.5 });
	});

	it('puts a drop shadow beneath the shape it belongs to', async () => {
		const { page, notes, content } = await draw(
			`<svg ${SVG}><defs><filter id="s"><feDropShadow dx="3" dy="4" stdDeviation="2" flood-color="#000000" flood-opacity="0.4"/></filter></defs>` +
				`<rect x="20" y="20" width="50" height="30" fill="#ff0000" filter="url(#s)"/></svg>`,
		);
		expect(notes).not.toContain('filterDropped');
		expect(content.indexOf('/Im0 Do')).toBeLessThan(content.indexOf('1 0 0 rg'));
		expect(page.alphas.get('GS0')).toEqual({ fill: 0.4, stroke: 0.4 });
		// The silhouette is shifted by the shadow offset: its box starts at 20 + 3 - 3 deviations.
		expect(content).toMatch(/\n[\d.]+ 0 0 -[\d.]+ 17 \S+ cm\n\/Im0 Do/);
	});

	it('paints text shadows as soft copies and keeps reporting shadows it cannot blur', async () => {
		const { notes, content } = await draw(
			`<svg ${SVG}><defs><filter id="f"><feGaussianBlur stdDeviation="2"/></filter></defs>` +
				`<g data-shadow="" filter="url(#f)"><text x="0" y="10" fill="#000000">Hi</text></g></svg>`,
		);
		expect(notes).not.toContain('shadowBlurDropped');
		expect(content).toContain('/ActualText ()');
		expect(content).toContain('(Hi) Tj');
		const mixed = await draw(
			`<svg ${SVG}><defs><filter id="f"><feGaussianBlur stdDeviation="2"/></filter></defs>` +
				`<g data-shadow="" filter="url(#f)"><rect width="5" height="5" fill="#000000"/><rect width="5" height="5" fill="#ff0000"/></g></svg>`,
		);
		expect(mixed.notes).toContain('shadowBlurDropped');
	});

	it('reads only the filters it can paint', async () => {
		const filter = async (inner: string) =>
			childElements(
				(await parseSvgMarkup(`<svg ${SVG}><filter id="f">${inner}</filter></svg>`))!,
			)[0];
		expect(readShadowFilter(await filter('<feGaussianBlur stdDeviation="3 1"/>'))).toEqual({
			kind: 'blur',
			sigma: 3,
		});
		expect(
			readShadowFilter(await filter('<feGaussianBlur in="SourceAlpha" stdDeviation="3"/>')),
		).toBe(undefined);
		expect(
			readShadowFilter(
				await filter('<feMorphology radius="2"/><feGaussianBlur stdDeviation="3"/>'),
			),
		).toBeUndefined();
		expect(readShadowFilter(await filter('<feDropShadow stdDeviation="1"/>'))).toMatchObject({
			kind: 'drop',
			dx: 2,
			dy: 2,
			color: { r: 0, g: 0, b: 0, a: 1 },
		});
	});
});
