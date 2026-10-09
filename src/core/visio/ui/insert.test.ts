import { describe, expect, it } from 'vitest';
import type { VisioDocument, VisioPage, VisioShape } from '../model';
import type { VisioHyperlink } from '../shape-metadata';
import {
	normalizeVisioHyperlinkAddress,
	visioFollowTarget,
	visioPictureInsertCommand,
} from './insert';

const page = (patch: Partial<VisioPage> = {}): VisioPage =>
	({
		id: '0',
		name: 'Page-1',
		width: 8.5,
		height: 11,
		shapes: [],
		connectors: [],
		isBackground: false,
		...patch,
	}) as VisioPage;
const image = new Uint8Array([1]);

describe('picture placement', () => {
	it('centres the natural 96 DPI size and keeps the aspect ratio when scaling to the page', () => {
		expect(
			visioPictureInsertCommand(page(), { pixelWidth: 192, pixelHeight: 96 }, image),
		).toMatchObject({
			type: 'insert-picture',
			shapeId: '1',
			x: 4.25,
			y: 5.5,
			width: 2,
			height: 1,
		});
		const big = visioPictureInsertCommand(page(), { pixelWidth: 9600, pixelHeight: 4800 }, image);
		expect(big.width).toBeCloseTo(6.8);
		expect(big.width / big.height).toBeCloseTo(2);
	});
	it('converts to drawing inches on scaled pages and allocates a free shape ID', () => {
		const shapes = [{ id: '7', children: [] } as unknown as VisioShape];
		const command = visioPictureInsertCommand(
			page({ drawingToPageScale: 0.5, shapes }),
			{ pixelWidth: 96, pixelHeight: 96 },
			image,
		);
		expect(command).toMatchObject({ shapeId: '8', x: 8.5, y: 11, width: 2, height: 2 });
	});
});

describe('link helpers', () => {
	it('normalises bare hosts and e-mail addresses only', () => {
		expect(normalizeVisioHyperlinkAddress(' example.com/a ')).toBe('https://example.com/a');
		expect(normalizeVisioHyperlinkAddress('me@example.com')).toBe('mailto:me@example.com');
		expect(normalizeVisioHyperlinkAddress('http://x.test')).toBe('http://x.test');
		expect(normalizeVisioHyperlinkAddress('notes.txt')).toBe('notes.txt');
		expect(normalizeVisioHyperlinkAddress('')).toBe('');
	});
	it('follows the default visible link to a URL or an existing page, never an unresolved one', () => {
		const link = (target: VisioHyperlink['target'], extra: Partial<VisioHyperlink> = {}) =>
			({ id: '1', name: 'Row_1', target, ...extra }) as VisioHyperlink;
		const document = { pages: [page(), page({ id: '1', name: 'Details' })] } as VisioDocument;
		const shape = (hyperlinks: VisioHyperlink[]) => ({ hyperlinks }) as VisioShape;
		expect(
			visioFollowTarget(
				document,
				shape([
					link({ kind: 'external', href: 'https://a.test/' }),
					link({ kind: 'internal', subAddress: 'details/Sheet.1' }, { default: true }),
				]),
			),
		).toEqual({ kind: 'page', index: 1 });
		expect(
			visioFollowTarget(document, shape([link({ kind: 'external', href: 'https://a.test/' })])),
		).toEqual({ kind: 'external', href: 'https://a.test/' });
		expect(
			visioFollowTarget(document, shape([link({ kind: 'unresolved', reason: 'x' })])),
		).toBeUndefined();
		expect(
			visioFollowTarget(
				document,
				shape([link({ kind: 'external', href: 'https://a.test/' }, { invisible: true })]),
			),
		).toBeUndefined();
		expect(
			visioFollowTarget(document, shape([link({ kind: 'internal', subAddress: 'Nope' })])),
		).toBeUndefined();
	});
});
