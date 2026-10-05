import { describe, it, expect } from 'vitest';
import { demoDocument } from './demo-document.js';
import { shapeDetails, safeExternalHref, selectedShape } from './shape-inspector.js';
describe('read-only shape inspector', () => {
	it('renders cached values literally and honors invisible fields', () => {
		const shape = structuredClone(demoDocument.pages[0]!.shapes[0]!);
		shape.shapeData = [
			{
				id: '0',
				name: 'Owner',
				label: 'Responsible person',
				type: 0,
				valueKind: 'string',
				value: '<script>text</script>',
			},
			{
				id: '1',
				name: 'Hidden',
				type: 0,
				valueKind: 'string',
				value: 'private field',
				invisible: true,
			},
		];
		const host = document.createElement('div');
		host.append(shapeDetails(shape));
		expect(host.textContent).toContain('Responsible person');
		expect(host.textContent).toContain('<script>text</script>');
		expect(host.querySelector('script')).toBeNull();
		expect(host.textContent).not.toContain('private field');
	});
	it('creates only safe external anchors and leaves internal targets inert', () => {
		const shape = structuredClone(demoDocument.pages[0]!.shapes[0]!);
		shape.hyperlinks = [
			{ id: '0', name: 'Safe', target: { kind: 'external', href: 'https://example.com/path' } },
			{ id: '1', name: 'Unsafe', target: { kind: 'external', href: 'javascript:alert(1)' } },
			{ id: '2', name: 'Inside', target: { kind: 'internal', subAddress: 'Page-2/Sheet.4' } },
		];
		const host = document.createElement('div');
		host.append(shapeDetails(shape));
		expect(host.querySelectorAll('a')).toHaveLength(1);
		expect(host.querySelector('a')?.rel).toBe('noopener noreferrer');
		expect(host.textContent).toContain('internal target: Page-2/Sheet.4');
	});
	it.each([
		'javascript:alert(1)',
		'data:text/html,x',
		'file:///etc/passwd',
		'https://user:pass@example.com/',
		'https://example.com/%0a',
		'https://example.com/%zz',
		'https:\\example.com',
	])('rejects unsafe programmatic href %s', (value) => {
		expect(safeExternalHref(value)).toBeUndefined();
	});
	it('resolves page-scoped selection against the correct tree', () => {
		const model = structuredClone(demoDocument);
		model.pages[1]!.shapes[0]!.id = 's1';
		expect(selectedShape(model, { id: 's1', name: '', pageId: '2' }, 0)).toBe(
			model.pages[1]!.shapes[0],
		);
	});
});

describe('metadata display budgets', () => {
	it('bounds a property preview without changing the source values', () => {
		const shape = structuredClone(demoDocument.pages[0]!.shapes[0]!);
		shape.shapeData = Array.from({ length: 250 }, (_, index) => ({
			id: String(index),
			name: 'Field',
			type: 0,
			valueKind: 'string' as const,
			value: 'x'.repeat(8192),
		}));
		const host = document.createElement('div');
		host.append(shapeDetails(shape));
		expect(host.textContent!.length).toBeLessThan(105_000);
		expect(host.textContent).toContain('shortened');
		expect(shape.shapeData[0]!.value).toHaveLength(8192);
	});
	it('does not silently repair invalid HTTP authority syntax', () => {
		expect(safeExternalHref('https:///example.com')).toBeUndefined();
		expect(safeExternalHref('https://@example.com')).toBeUndefined();
	});
});
