import { describe, expect, it } from 'vitest';

import {
	isAllowedResourceUrl,
	isBrowserOpenableDataUrl,
	isBrowserOpenableType,
	normalizeMimeType,
	parseDataUrlHeader,
} from './safe-resource-url';

describe('parseDataUrlHeader', () => {
	it('reads the type and encoding from the URL', () => {
		expect(parseDataUrlHeader('data:image/png;base64,iVBOR')).toEqual({
			mime: 'image/png',
			base64: true,
		});
		expect(parseDataUrlHeader('DATA:Text/Plain;charset=utf-8,hi')).toEqual({
			mime: 'text/plain',
			base64: false,
		});
		expect(parseDataUrlHeader('data:,hi')).toEqual({ mime: 'text/plain', base64: false });
		expect(parseDataUrlHeader('data:image/svg+xml;utf8,<svg/>')?.mime).toBe('image/svg+xml');
	});

	it('refuses anything that is not a well-formed data URL', () => {
		for (const url of [
			'https://example.com/a.png',
			'data:text/html<script>,x',
			'data: text/html,x',
			'data:texthtml,x',
			'data:image/png;base64',
			`data:image/png;${'a'.repeat(300)},x`,
			42,
		]) {
			expect(parseDataUrlHeader(url)).toBeUndefined();
		}
	});
});

describe('isBrowserOpenableDataUrl', () => {
	it('opens inert types and nothing that can hold markup or script', () => {
		expect(isBrowserOpenableDataUrl('data:application/pdf;base64,JVBER')).toBe(true);
		expect(isBrowserOpenableDataUrl('data:image/png;base64,iVBOR')).toBe(true);
		for (const url of [
			'data:text/html,<script>1</script>',
			'data:image/svg+xml,<svg/>',
			'data:application/xhtml+xml,x',
			'data:text/javascript,1',
			'javascript:alert(1)',
			'https://example.com/a.pdf',
		]) {
			expect(isBrowserOpenableDataUrl(url)).toBe(false);
		}
	});

	it('judges a bare type the same way', () => {
		expect(isBrowserOpenableType('application/PDF; charset=x')).toBe(true);
		expect(isBrowserOpenableType('text/html')).toBe(false);
		expect(isBrowserOpenableType(undefined)).toBe(false);
	});
});

describe('isAllowedResourceUrl', () => {
	it('accepts data URLs of the kind the field holds', () => {
		expect(isAllowedResourceUrl('data:image/png;base64,x', { kind: 'image' })).toBe(true);
		expect(isAllowedResourceUrl('data:video/mp4;base64,x', { kind: 'media' })).toBe(true);
		expect(isAllowedResourceUrl('data:model/gltf-binary;base64,x', { kind: 'model' })).toBe(true);
		expect(
			isAllowedResourceUrl('data:application/vnd.ms-excel.sheet.macroEnabled.12;base64,x', {
				kind: 'file',
			}),
		).toBe(true);
		expect(isAllowedResourceUrl('data:text/html,x', { kind: 'image' })).toBe(false);
		expect(isAllowedResourceUrl('data:text/html,x', { kind: 'media' })).toBe(false);
	});

	it('accepts web URLs and package paths only where the field allows them', () => {
		expect(isAllowedResourceUrl('https://example.com/a.png', { kind: 'image', web: true })).toBe(
			true,
		);
		expect(isAllowedResourceUrl('https://example.com/a.png', { kind: 'image' })).toBe(false);
		expect(isAllowedResourceUrl('ppt/media/image1.png', { kind: 'image', path: true })).toBe(true);
		expect(isAllowedResourceUrl('ppt/media/image1.png', { kind: 'image' })).toBe(false);
	});

	it('never accepts script, local or disguised schemes', () => {
		const rule = { kind: 'file', web: true, path: true } as const;
		for (const url of [
			'javascript:alert(1)',
			'JaVaScRiPt:alert(1)',
			'java\tscript:alert(1)',
			'vbscript:x',
			'blob:https://example.com/1',
			'file:///C:/a.png',
			'',
			undefined,
		]) {
			expect(isAllowedResourceUrl(url, rule)).toBe(false);
		}
	});
});

describe('normalizeMimeType', () => {
	it('lower-cases a well-formed type and refuses anything else', () => {
		expect(normalizeMimeType('Image/PNG')).toBe('image/png');
		expect(normalizeMimeType('text/html"><b>')).toBeUndefined();
		expect(normalizeMimeType('text')).toBeUndefined();
	});
});
